using Carubbi.FileDownloaderScheduler.Configuration;
using Carubbi.FileDownloaderScheduler.PluginInterfaces;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Timers;
using Timer = System.Timers.Timer;

namespace Carubbi.FileDownloaderScheduler
{
    public class Program
    {
        private static StringCollection _paths;
        private static int _minutesCycle;
        private static string _targetPath;
        private static Timer _timer;
        private static List<IFileDownloaderSchedulerPlugin> _plugins;
        private static string _prefixFileName;
        private static string _sufixFileName;
        private static bool _isOnline = true;

        private static void ReadConfigs()
        {
            try
            {
                _prefixFileName = ConfigurationManager.AppSettings["prefixFileName"];
                _sufixFileName = ConfigurationManager.AppSettings["sufixFileName"];
                _isOnline = Convert.ToBoolean(ConfigurationManager.AppSettings["isOnline"]);

                Console.WriteLine("{0} - Lendo configurações", DateTime.Now);
                Console.WriteLine("1.) Urls a monitorar", DateTime.Now);
                _paths = UrlSourcesConfig.Paths;
                var urlsCount = 0;
                foreach (var path in _paths)
                    Console.WriteLine("- Url {0}: {1} ", ++urlsCount, path);

                _minutesCycle = Convert.ToInt32(ConfigurationManager.AppSettings["minutesCycle"]);
                Console.WriteLine("2.) Ciclo de ativação: A cada {0} minuto(s)", _minutesCycle);

                _targetPath = ConfigurationManager.AppSettings["targetPath"];
                Console.WriteLine("3.) Caminho de saida: {0}", _targetPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ocorreu um erro ao ler as configurações: {0}", ex.Message);
            }
        }

        private static void Main(string[] args)
        {
            Console.WriteLine("{0} - File Downloader Scheduler iniciado", DateTime.Now);

            ReadConfigs();
            if (_isOnline)
            {
                _timer = new Timer(10000); // * 10000 * 6);
                _plugins = LoadPlugins();
                _timer.Elapsed += _timer_Elapsed;
                _timer.Start();


                while (true)
                    Thread.Sleep(100000);
            }

            _plugins = LoadPlugins();
            _timer_Elapsed(null, null);
        }

        public static List<KeyValuePair<string, Stream>> ExecuteDownloads(
            out List<KeyValuePair<string, Exception>> erros)
        {
            erros = new List<KeyValuePair<string, Exception>>();
            var arquivos = new List<KeyValuePair<string, Stream>>();
            foreach (var path in _paths)
                try
                {
                    var request = WebRequest.Create(path);
                    if (request is HttpWebRequest)
                        request.UseDefaultCredentials = true;
                    var response = request.GetResponse();
                    var fileContent = response.GetResponseStream();
                    arquivos.Add(new KeyValuePair<string, Stream>(path, fileContent));
                    Console.WriteLine("{0} - Download efetuado de {1}", DateTime.Now, path);
                }
                catch (Exception ex)
                {
                    erros.Add(new KeyValuePair<string, Exception>("Erro ao tentar realizar download", ex));
                }

            return arquivos;
        }

        private static void _timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Console.WriteLine("------------------------------------------------------------");
            Console.WriteLine("{0} - Ciclo de execução iniciado", DateTime.Now);

            if (_isOnline)
                _timer.Stop();

            var arquivos = ExecuteDownloads(out var erros);
            ProcessarPlugins(ref arquivos, erros);
            GravarSaidas(ref arquivos, erros);

            foreach (var erro in erros) Console.WriteLine(string.Concat(erro.Key, ":", erro.Value.Message));

            if (_isOnline)
                _timer.Start();
        }

        private static void ProcessarPlugins(ref List<KeyValuePair<string, Stream>> arquivos,
            ICollection<KeyValuePair<string, Exception>> erros)
        {
            try
            {
                var arquivosProcessados = new List<KeyValuePair<string, Stream>>();
                var arquivosOutput = new List<KeyValuePair<string, Stream>>();

                if (_plugins.Count > 0)
                {
                    foreach (var arquivo in arquivos)
                    {
                        Console.WriteLine("Arquivo {0}: {1} plugins a processar...", arquivo.Key, _plugins.Count);


                        foreach (var plugin in _plugins)
                        {
                            Console.WriteLine("Executando plugin {0}", plugin.Name);
                            arquivosOutput = plugin.Process(arquivo);
                        }

                        if (arquivosOutput != null && arquivosOutput.Count > 0)
                            arquivosProcessados.AddRange(arquivosOutput);
                        else
                            arquivosProcessados.Add(arquivo);
                    }

                    arquivos = arquivosProcessados;
                }
            }
            catch (Exception ex)
            {
                erros.Add(new KeyValuePair<string, Exception>("Erro ao tentar processar plugins", ex));
            }
        }

        private static List<IFileDownloaderSchedulerPlugin> LoadPlugins()
        {
            var plugins = new List<IFileDownloaderSchedulerPlugin>();
            var pluginsDirectory = string.Concat(Environment.CurrentDirectory, @"\plugins\");
            var pluginFileNames = Directory.GetFiles(pluginsDirectory);

            try
            {
                foreach (var pluginFileName in pluginFileNames)
                {
                    var path = Path.Combine(pluginsDirectory, pluginFileName);
                    var plugin = Assembly.LoadFrom(path);
                    try
                    {
                        var types = plugin.GetTypes();
                        plugins.AddRange(from t 
                            in types
                            let interfaceTypes = t.GetInterfaces()
                            from interfaceType 
                                in interfaceTypes
                            where interfaceType == typeof(IFileDownloaderSchedulerPlugin)
                            select (IFileDownloaderSchedulerPlugin) Activator.CreateInstance(t));
                    }
                    catch (ReflectionTypeLoadException)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Plugin não pode ser carregado", ex.Message);
            }

            return plugins;
        }


        private static void GravarSaidas(ref List<KeyValuePair<string, Stream>> arquivos,
            ICollection<KeyValuePair<string, Exception>> erros)
        {
            foreach (var arquivo in arquivos)
                try
                {
                    var path = string.Empty;
                    var fileName = string.Empty;
                    path = arquivo.Key.Contains(Environment.CurrentDirectory) ? arquivo.Key.Remove(0, Environment.CurrentDirectory.Length + 1) : arquivo.Key;

                    fileName =
                        $"{ResolvePatternFileName(_prefixFileName)}{Path.Combine(Path.GetDirectoryName(path), Path.GetFileNameWithoutExtension(path))}{ResolvePatternFileName(_sufixFileName)}{Path.GetExtension(path)}";


                    var target = string.Concat(_targetPath, fileName);
                    if (!Directory.Exists(Path.GetDirectoryName(target)))
                        Directory.CreateDirectory(Path.GetDirectoryName(target));

                    using (Stream file = File.OpenWrite(target))
                    {
                        arquivo.Value.CopyTo(file); 
                    }

                    arquivo.Value.Close();
                    arquivo.Value.Dispose();

                    Console.WriteLine("{0} - Arquivo Gravado em {1}", DateTime.Now,
                        string.Concat(_targetPath, fileName));
                }
                catch (Exception ex)
                {
                    erros.Add(new KeyValuePair<string, Exception>("Erro ao tentar Gravar arquivo(s) de saida", ex));
                }

            arquivos.Clear();
        }


        public static string ResolvePatternFileName(string pattern)
        {
            return pattern.Replace("[YEAR]", DateTime.Today.ToString("yyyy"))
                .Replace("[MONTH]", DateTime.Today.ToString("MM")).Replace("[DAY]", DateTime.Today.ToString("dd"))
                .Replace("[TIME]", DateTime.Now.ToString("HHmmss"));
        }
    }
}