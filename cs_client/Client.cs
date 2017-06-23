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
using System.Net;
using System.Net.Sockets;

namespace cs_client
{
    class Client
    {
        public string ip;
        public string response;
        public bool debug = true;

        public Client() { }

        public Client(string ipAddress)
        {
            ip = ipAddress;
        }

        public bool ScanProcess(string pidOrName)
        {
            if (!SendRecv("attach:" + pidOrName + "\x0d")) return false;
            if (response.IndexOf("fail:") >= 0) return false;
            return true;
        }

        public bool ResolveExport(string apiOrAddress)
        {
            if (!SendRecv("resolve:" + apiOrAddress + "\x0d")) return false;
            if (response.Substring(0, 3) == "ok:")
            {
                response = response.Substring(3);
                if (response.IndexOf("Error:") >= 0) return false;
                response = response.Replace(" ", "");
            }
            return true;

        }


        private bool SendRecv(string data)
        {
            // Data buffer for incoming data.
            byte[] bytes = new byte[1024];
            response = "";

            if (ip.Length == 0)
            {
                response = "set ip address";
                return false;
            }

            // Connect to a remote device.
            try {
                // Establish the remote endpoint for the socket.
                IPHostEntry ipHostInfo = Dns.Resolve(ip);
                IPAddress ipAddress = ipHostInfo.AddressList[0];
                IPEndPoint remoteEP = new IPEndPoint(ipAddress,9000);

                // Create a TCP/IP  socket.
                Socket sender = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp );

                // Connect the socket to the remote endpoint. Catch any errors.
                try {
                    sender.Connect(remoteEP);

                    // Encode the data string into a byte array.
                    byte[] msg = Encoding.ASCII.GetBytes(data);

                    // Send the data through the socket.
                    int bytesSent = sender.Send(msg);

                    // Receive the response from the remote device.
                    int bytesRec = sender.Receive(bytes);
                    response = Encoding.ASCII.GetString(bytes,0,bytesRec);

                    // Release the socket.
                    sender.Shutdown(SocketShutdown.Both);
                    sender.Close();
                    
                } catch (ArgumentNullException ane) {
                    response = "ArgumentNullException :" + ane.ToString() ;
                    return false;
                } catch (SocketException se) {
                    if(response.Length==0) response = "SocketException : " + se.ToString() ;
                    return false;
                } catch (Exception e) {
                    if (response.Length == 0) response = "Unexpected exception : " + e.ToString();
                    return false;
                }

            } catch (Exception e) {
                response =  e.ToString() ;
                return false;
            }

            return true;
        }

    }
}
