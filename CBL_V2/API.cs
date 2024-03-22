﻿using Crestron.SimplSharp;
using Crestron.SimplSharp.Net;
using Crestron.SimplSharp.WebScripting;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;



namespace CBL
{
    
    public class API
    {
        private string _filename = "";
        public static Configuration Setting;  //Static?  YES!  This makes sure we have only ONE as a static member meaning it can exist only once.

        private HttpCwsRoute myRoute;
        private HttpCwsServer myServer;
        public API(string Filename)
        {
            Setting = new Configuration();                              // Create an instance of our Configuration class that will
                                                                        // be available as ClassName.Setting when we instantiate the
                                                                        // config class

           /* _filename = CheckFilename(Filename);                        // Is the filename valid and add in a path if one is not
            CWSDebug.Msg($"ProgID = {InitialParametersClass.ProgramIDTag}");
            CWSDebug.Msg($"RoomID = {InitialParametersClass.RoomId}");
            CWSDebug.Msg($"RoomName = {InitialParametersClass.RoomName}");
            CWSDebug.Msg($"ControllerPrompt = {InitialParametersClass.ControllerPromptName}");
            CWSDebug.Msg($"ProgramDirectory = {InitialParametersClass.ProgramDirectory}");
            Cws("app", "settings"); // Create the CWS server for the settings*/
            // /cws/app/settings
        }


        public void Cws(string path,string route)
        {
            myServer = new HttpCwsServer("/" + path);
            myRoute = new HttpCwsRoute(route + "/{REQUEST}");
            myRoute.RouteHandler = new ConfigRequestHandler();
            myServer.AddRoute(myRoute);
            if (myServer.Register())
            {

            }
        }
        public class ConfigRequestHandler : IHttpCwsHandler
        {

            private HttpCwsContext _context;
            public void ProcessRequest(HttpCwsContext context)
            {
                _context = context;

                var requestMethod = _context.Request.HttpMethod;
                var requestContents = "";

                using (var myreader = new Crestron.SimplSharp.CrestronIO.StreamReader(_context.Request.InputStream))
                {
                    requestContents = myreader.ReadToEnd().Trim();
                }
                switch(requestMethod)
                {
                    case "GET":
                        {
                            if(_context.Request.RouteData.Values.ContainsKey("REQUEST"))
                            {
                                if(_context.Request.RouteData.Values["REQUEST"].ToString().ToLower() == "endpoints")
                                {
                                    string response = JsonConvert.SerializeObject(API.Setting.Endpoints);
                                    GenerateResponseHeader();
                                    _context.Response.Write(response, true);
                                }
                                else if (_context.Request.RouteData.Values["REQUEST"].ToString().ToLower() == "all")
                                {
                                    string response = JsonConvert.SerializeObject(API.Setting);
                                    GenerateResponseHeader();
                                    _context.Response.Write(response, true);
                                }
                                else
                                {
                                    _context.Response.Write("Error in the request", true);
                                }
                            }
                            return;
                        }
                    case "POST":
                        {

                            if (_context.Request.RouteData.Values.ContainsKey("REQUEST"))
                            {
                                if (_context.Request.RouteData.Values["REQUEST"].ToString().ToLower() == "endpoints")
                                {

                                    using (var myreader = new Crestron.SimplSharp.CrestronIO.StreamReader(_context.Request.InputStream))
                                    {
                                        API.Setting.Endpoints = JsonConvert.DeserializeObject<List<NVX>>(requestContents);
                                    }
                                        _context.Response.Write("Endpoint OK", true);
                                }
                                else if (_context.Request.RouteData.Values["REQUEST"].ToString().ToLower() == "all")
                                {
                                    using (var myreader = new Crestron.SimplSharp.CrestronIO.StreamReader(_context.Request.InputStream))
                                    {
                                        API.Setting.Endpoints = JsonConvert.DeserializeObject<Configuration>(requestContents);
                                    }
                                    _context.Response.Write("OK", true);
                                }
                                else
                                {
                                    _context.Response.Write("Error in the request", true);
                                }
                            }






                            return;
                        }
                }
            }

            private void GenerateResponseHeader()
            {
                _context.Response.StatusCode = 200;
                _context.Response.ContentType = "text/plain";
            }


        }

        public class NVX
        {
            public string Address = "";
            public string Name = "";
        }
        public class Configuration
        {

            public string IPAddress = "";
            public ushort Port = 0;


            public List<NVX> Endpoints;

            public Configuration()
            {
                Endpoints = new List<NVX>();
            }
        
        }






    }
}
