/*
	Contributed by FireEye FLARE Team
	Author:  David Zimmer <david.zimmer@fireeye.com>, <dzzie@yahoo.com>
	Copyright (C) 2017 FireEye, Inc. All Rights Reserved.
	License: GPL
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace cs_client
{
    class Program
    {
        static void Main(string[] args)
        {
            Client c = new Client("192.168.0.10");
            Console.WriteLine("Server: " + c.ip);

            bool scanOk = c.ScanProcess("explorer");
            Console.WriteLine(c.response);
            if(!scanOk) goto cleanup;

            string[] apis = {"GetProcAddress", "ntdll.atoi", "msvcrt!atoi", "0x7c42837"};

            foreach(string api in apis){
                bool r = c.ResolveExport(api);
                Console.WriteLine("{0,-15}= {1,-5}   val = '{2}'", api, r, c.response );
            }

            

        cleanup:
            Console.WriteLine("\n\nPress any key to exit...");
            Console.ReadKey();
        }
    }
}
