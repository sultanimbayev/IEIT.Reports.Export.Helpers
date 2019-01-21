using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Usage.Interfaces
{
    public interface ICreateFile
    {
        /// <summary>
        /// Returns created file path
        /// </summary>
        /// <returns></returns>
        string CreateOne();
    }
}
