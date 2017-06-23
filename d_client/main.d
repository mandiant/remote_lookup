import std.stdio;
import std.socket;
import std.socketstream;
import std.stream;

string domain = "192.168.0.10";
ushort port=9000;

void sendRecv(string msg){
	Socket sock = new TcpSocket(new InternetAddress(domain, port));
    Stream ss = new SocketStream(sock);
	ss.writeString(msg ~ "\r");
    while (!ss.eof())
    {
        auto line = ss.readLine();
        writeln(line);
    }
	sock.close();
}

int main(string[] argv)
{
	writefln("Connecting to %s...", domain);
	sendRecv("attach:explorer");
	sendRecv("resolve:getprocaddress");
    return 0;
}
