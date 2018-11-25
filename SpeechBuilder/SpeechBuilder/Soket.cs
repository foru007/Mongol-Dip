using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using System.IO;
using System.Windows.Forms;

namespace SpeechBuilder
{
    public class Soket
    {
        public String mm(String ss)
        {            
            try
            {                               
                //Console.WriteLine(ss);     
                TcpClient tcp = new TcpClient("localhost", 2085);                
                NetworkStream ns = tcp.GetStream();
                //NetworkStream ns1 = tcp1.GetStream();
                StreamWriter sw = new StreamWriter(ns);
                StreamReader sr = new StreamReader(ns);
                sw.WriteLine(ss);
                sw.AutoFlush = true;                
                String rs=sr.ReadLine().ToString();
                Console.WriteLine(rs);
                sw.Close();
                sr.Close();
                ns.Close();
                return rs;
            }
            catch (Exception ex) { 
               // MessageBox.Show("1: "+ ex.ToString()); 
            }
            return null;
        }
        public void mmm(String ss)
        {
            try
            {
                TcpClient tcp1 = new TcpClient("localhost", 2087);
                NetworkStream ns = tcp1.GetStream();
                StreamWriter sw = new StreamWriter(ns);
                sw.WriteLine(ss);
                sw.AutoFlush = true;
                sw.Close();
                ns.Close();               
            }
            catch (Exception ex) { 
               // MessageBox.Show("2: " + ex.ToString());  
            }
        }
         
    }
}
