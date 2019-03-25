using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace ExcelAddIn.Access {
    public static class Configuration
    {
        /// <summary>Función para invocar el parámetro del App.config.
        /// <para>Invocar la llave del archivo App.config. Referencia: <see cref="_getConfig(string)"/></para>
        /// <seealso cref="_getConfig(string)"/>
        /// </summary>
        static string _getConfig(string _Key) => ConfigurationManager.AppSettings[_Key];
        /// <summary>Función para desencriptar el parámetro del App.config.
        /// <para>Desencritar la llave del archivo App.config. Referencia: <see cref="_unEncrypt(string)"/></para>
        /// <seealso cref="_unEncrypt(string)"/>
        /// </summary>
        static string _unEncrypt(string _Value) => Encoding.UTF8.GetString(Convert.FromBase64String(_Value));
        /// <summary>Función para obtener la cadena de Conexión del archivo App.config.
        /// <para>Desencritar la llave de Conexión del archivo App.config. Referencia: <see cref="ConnectionString"/></para>
        /// <seealso cref="ConnectionString"/>
        /// </summary>
        public static string ConnectionString => _unEncrypt(_getConfig("VAL0"));
        /// <summary>Función para obtener el nombre del Servidor de SQL Server del archivo App.config.
        /// <para>Desencritar el nombre del Servidor del archivo App.config. Referencia: <see cref="Server"/></para>
        /// <seealso cref="Server"/>
        /// </summary>
        public static string Server => _unEncrypt(_getConfig("VAL1"));
        /// <summary>Función para obtener el nombre de Base de Datos de SQL Server del archivo App.config.
        /// <para>Desencritar el nombre de Base de Datos del Servidor en el archivo App.config. Referencia: <see cref="DataBase"/></para>
        /// <seealso cref="DataBase"/>
        /// </summary>
        public static string DataBase => _unEncrypt(_getConfig("VAL2"));
        /// <summary>Función para obtener el Usuario de Base de Datos de SQL Server del archivo App.config.
        /// <para>Desencritar el Usuario de Base de Datos del Servidor en el archivo App.config. Referencia: <see cref="User"/></para>
        /// <seealso cref="User"/>
        /// </summary>
        public static string User => _unEncrypt(_getConfig("VAL3"));
        /// <summary>Función para obtener el Password de Base de Datos de SQL Server del archivo App.config.
        /// <para>Desencritar el Password de Base de Datos del Servidor en el archivo App.config. Referencia: <see cref="Password"/></para>
        /// <seealso cref="Password"/>
        /// </summary>
        public static string Password => _unEncrypt(_getConfig("VAL4"));
        /// <summary>Función para obtener el Tiempo de Espera para la Conexión de Base de Datos de SQL Server del archivo App.config.
        /// <para>Desencritar el Tiempo de Espera para la Conexión de Base de Datos del Servidor en el archivo App.config. Referencia: <see cref="TimeOut"/></para>
        /// <seealso cref="TimeOut"/>
        /// </summary>
        public static int TimeOut => int.Parse(_unEncrypt(_getConfig("VAL5")));
        /// <summary>Función para obtener el Path de los Archivos de Instalación en el archivo App.config.
        /// <para>Desencritar el Path de los Archivos de Instalación en el archivo App.config. Referencia: <see cref="Path"/></para>
        /// <seealso cref="Path"/>
        /// </summary>
        public static string Path => _getConfig("VAL6");
        /// <summary>Función para obtener el Password del Archivo de Excel en el archivo App.config.
        /// <para>Desencritar el Password del Archivo de Excel en el archivo App.config. Referencia: <see cref="PwsExcel"/></para>
        /// <seealso cref="PwsExcel"/>
        /// </summary>
        public static string PwsExcel => _unEncrypt(_getConfig("VAL7"));
        /// <summary>Función para obtener la URL de Conexión de los Archivos Json en el archivo App.config.
        /// <para>Desencritar la URL de Conexión en el archivo App.config. Referencia: <see cref="UrlConnection"/></para>
        /// <seealso cref="UrlConnection"/>
        /// </summary>
        public static string UrlConnection => _getConfig("VAL8");
        /// <summary>
        /// Función para obtener la cadena de conexión para el prellenado de la plantilla
        /// </summary>
        public static string ConnectionStringPrellenado => _unEncrypt(_getConfig("VAL9"));
    }
}