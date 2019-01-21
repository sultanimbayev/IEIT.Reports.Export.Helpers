using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Usage.Interfaces;

namespace Usage
{
    public class ExamplesFabric
    {
        public T GetExample<T>(string name) where T : class
        {
            var assembly = Assembly.GetExecutingAssembly();
            var type = typeof(T);
            var rgx = new Regex($"{name}(Example)?$");
            var fileCreatorType = assembly.GetTypes().FirstOrDefault(t => !t.IsAbstract && !t.IsInterface && type.IsAssignableFrom(t) && rgx.IsMatch(t.Name));
            if (fileCreatorType == null)
            {
                return null;
            }
            var fileCreator = Activator.CreateInstance(fileCreatorType);
            return fileCreator as T;
        }
    }
}
