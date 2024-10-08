﻿


// -------------------------------------------------------------------------------------------

// Copyright (c) 2024 Federico Turco

// Permission is hereby granted, free of charge, to any person
// obtaining a copy of this software and associated documentation
// files (the "Software"), to deal in the Software without
// restriction, including without limitation the rights to use,
// copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the
// Software is furnished to do so, subject to the following
// conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
// OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
// HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
// WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
// OTHER DEALINGS IN THE SOFTWARE.

// -------------------------------------------------------------------------------------------

// NB: I file in pdf accessibili dal menu info sono di proprietà dei rispettivi autori

// -------------------------------------------------------------------------------------------



using System;
//using System.Drawing;
using System.Collections.Concurrent;
using System.IO;
//Porta seriale
using System.IO.Ports;
//Comandi apri/chiudi console

//Libreria JSON

//Classecon funzioni di conversione DEC-HEX
using System.Net.Security;
//Process.

//Sockets
using System.Net.Sockets;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Authentication;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
//Threading per server ModBus TCP
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ModBusMaster_Chicco
{
    public class ModBus_Def
    {
        public int TYPE_RTU = 10;        // Mode RTU
        public int TYPE_ASCII = 20;      // Mode ASCII (not supported now)
        public int TYPE_TCP_REOPEN = 30; // Open the connection on every request modbus
        public int TYPE_TCP_SOCK = 31;   // Open the conection only al first time
        public int TYPE_TCP_SECURE = 32;   // Open the conection with TLS ecnryption
    }

    public class ModBus_Chicco
    {
        //-----VARIABILI-OGGETTO------
        public bool ClientActive = false;

        bool ordineTextBoxLog = true;   // true -> in cima, false -> in fondo

        // RTU o ASCII
        SerialPort serialPort;

        // TCP
        TcpClient client;
        SslStream sslStream;
        NetworkStream stream;
        String ip_address;
        String port;
        String clientCertificateFile;
        String clientCertificatePassword;
        String clientKeyFile;

        public int type;        //"RTU", "ASCII", "TCP"
        public int tlsVersion;  // 0 -> 1.2, 1 -> 1.3, ...
        ModBus_Def def = new ModBus_Def();

        Border pictureBoxSending = new Border();
        Border pictureBoxReceiving = new Border();

        public FixedSizedQueue<String> log = new FixedSizedQueue<string>();
        public FixedSizedQueue<String> log2 = new FixedSizedQueue<string>();

        UInt16 queryCounter = 1;    //Conteggio richieste TCP per inserirli nel primo byte

        int buffer_dimension = 256; //Dimensione buffer per comandi invio/ricezione seriali/tcp

        bool TX_on = false;
        bool TX_set = false;
        bool RX_on = false;
        bool RX_set = false;

        bool threadTxRxIsRunning = false;
        Thread threadTxRx;

        /*
        01 Illegal Function
        02 Illegal Data Address
        03 Illegal Data Value
        04 Slave Device Failure
        05 Acknowledge
        06 Slave Device Busy
        07 Negative Acknowledge
        08 Memory Parity Error
        10 Gateway Path Unavailable
        11 Gateway Target Device Failed to Respond
         */
        public string[] ModbusErrorCodes = { "", "Illegal Function", "Illegal Data Address", "Illegal Data Value", "Slave Device Failure", "Acknowledge", "Slave Device Busy", "Negative Acknowledge", "Memory Parity Error", "", "Gateway Path Unavailable", "Gateway Target Device Failed to Respond" };


        // Lockfile stream TCP
        private object lockModbus = new object();

        // Statistic
        public uint Count_FC01 = 0;
        public uint Count_FC02 = 0;
        public uint Count_FC03 = 0;
        public uint Count_FC04 = 0;
        public uint Count_FC05 = 0;
        public uint Count_FC06 = 0;
        public uint Count_FC08 = 0;
        public uint Count_FC15 = 0;
        public uint Count_FC16 = 0;

        public ModBus_Chicco(
            SerialPort serialPort_, 
            String ip_address_, 
            String port_, 
            int type_, 
            Border pictureBoxSending_, 
            Border pictureBoxReceiving_,
            String clientCertificateFile_,
            String clientCertificatePassword_,
            String clientKeyFile_,
            int tlsVersion_)
        {
            //ClientActive = true;
            //Type: TCP, RTU, ASCII
            type = type_;

            //RTU/ASCII
            serialPort = serialPort_;

            //TCP
            ip_address = ip_address_;
            port = port_;
            clientCertificateFile = clientCertificateFile_;
            clientCertificatePassword = clientCertificatePassword_;
            clientKeyFile = clientKeyFile_;
            tlsVersion = tlsVersion_;

            //GRAFICA
            pictureBoxSending = pictureBoxSending_;
            pictureBoxReceiving = pictureBoxReceiving_;

            //Dimensione log locale
            log.Limit = 10000;
            log2.Limit = 10000;

            if (!threadTxRxIsRunning)
            {
                threadTxRx = new Thread(new ThreadStart(handleTxRxGui));
                //threadTxRx.IsBackground = true;
                threadTxRx.Start();
            }
        }

        public void handleTxRxGui()
        {
            threadTxRxIsRunning = true;

            long timeout = 250;

            DateTimeOffset now = DateTimeOffset.UtcNow;

            long TX_epoch = now.ToUnixTimeMilliseconds();
            long RX_epoch = now.ToUnixTimeMilliseconds();
            long millis = now.ToUnixTimeMilliseconds();

            while (threadTxRxIsRunning)
            {
                if (TX_set)
                {
                    if (!TX_on)
                    {
                        pictureBoxSending.Dispatcher.Invoke((Action)delegate
                        {
                            pictureBoxSending.Background = Brushes.Yellow;
                        });
                    }

                    TX_set = false;
                    TX_on = true;

                    now = DateTimeOffset.UtcNow;
                    TX_epoch = now.ToUnixTimeMilliseconds();

                    DoEvents();
                }

                if (RX_set)
                {
                    if (!RX_on)
                    {
                        pictureBoxReceiving.Dispatcher.Invoke((Action)delegate
                        {
                            pictureBoxReceiving.Background = Brushes.Yellow;
                        });
                    }

                    RX_set = false;
                    RX_on = true;

                    now = DateTimeOffset.UtcNow;
                    RX_epoch = now.ToUnixTimeMilliseconds();

                    DoEvents();
                }

                now = DateTimeOffset.UtcNow;
                millis = now.ToUnixTimeMilliseconds();

                if ((millis - TX_epoch) > timeout && TX_on)
                {
                    TX_on = false;

                    pictureBoxSending.Dispatcher.Invoke((Action)delegate
                    {
                        pictureBoxSending.Background = Brushes.LightGray;
                    });

                    DoEvents();
                }

                if ((millis - RX_epoch) > timeout && RX_on)
                {
                    RX_on = false;

                    pictureBoxReceiving.Dispatcher.Invoke((Action)delegate
                    {
                        pictureBoxReceiving.Background = Brushes.LightGray;
                    });

                    DoEvents();
                }

                Thread.Sleep(5);

                // debug
                //counter += 1;
                //Console.WriteLine(counter);
            }
        }

        public void open()
        {
            if (type == def.TYPE_TCP_SOCK)
            {
                client = new TcpClient(ip_address, int.Parse(port));
                stream = client.GetStream();
            }
            if (type == def.TYPE_TCP_REOPEN)
            {
                client = new TcpClient(ip_address, int.Parse(port));
                client.Close();
            }
            if (type == def.TYPE_TCP_SECURE)
            {
                SslProtocols tlsProtocol = SslProtocols.Tls12;
                if (tlsVersion == 0)
                    tlsProtocol = SslProtocols.Tls12;
                if (tlsVersion == 1)
                    tlsProtocol = SslProtocols.Tls13;

                // PFX
                var clientCertificate = new X509Certificate2(clientCertificateFile, clientCertificatePassword);
                var clientCertificateCollection = new X509CertificateCollection(new X509Certificate[] { clientCertificate });
                
                client = new TcpClient(ip_address, int.Parse(port));
                sslStream = new SslStream(client.GetStream(), false, App_CertificateValidation);
                sslStream.AuthenticateAsClient(
                    ip_address, 
                    clientCertificateCollection,
                    tlsProtocol, 
                    false);
            }

            ClientActive = true;
        }

        public void close()
        {
            threadTxRxIsRunning = false;
            ClientActive = false;

            if (type == def.TYPE_RTU)
            {
                try
                {
                    serialPort.Close();
                }
                catch { }
            }

            else if (type == def.TYPE_TCP_SOCK)
            {
                try
                {
                    client.Close();
                }
                catch { }
            }
            
            else if (type == def.TYPE_TCP_SECURE)
            {
                try
                {
                    sslStream.Close();
                }
                catch { }
                
                try
                {
                    client.Close();
                }
                catch { }
            }
        }

        bool App_CertificateValidation(Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            Console.WriteLine("sslPolicyErrors: {0}", sslPolicyErrors);

            if (sslPolicyErrors == SslPolicyErrors.None)
            {
                return true;
            }
            
            if (sslPolicyErrors == SslPolicyErrors.RemoteCertificateChainErrors)
            {
                return true;
            }

            if (sslPolicyErrors == (SslPolicyErrors.RemoteCertificateChainErrors | SslPolicyErrors.RemoteCertificateNameMismatch)) 
            {
                // Bypass error check RemoteName if certificate mismatch the ip server address
                Console.WriteLine("Check RemoteCertificateNameMismatch");
                return true; 
            }

            return false;
        }

        public bool checkModbusTcpResponse(byte[] buffer, int length, byte[] query)
        {
            if (length < 9)
            {
                Console.WriteLine("Invalid Length {0}", length);
                throw new ModbusException("ModbusProtocolError: Invalid Length: " + length.ToString());
            }

            if (query.Length > 1)
            {
                UInt16 transactionIdQuery = (UInt16)((buffer[0] << 8) + buffer[1]);
                UInt16 transactionIdResponse = (UInt16)((query[0] << 8) + query[1]);
                if (transactionIdQuery != transactionIdResponse)
                {
                    Console.WriteLine("Invalid TransactionID: {0:X}, expected: {1:X}", transactionIdResponse, transactionIdQuery);
                    throw new ModbusException("ModbusProtocolError: Invalid TransactionID: 0x" + transactionIdResponse.ToString("X").PadLeft(4, '0') + ", expected: 0x" + transactionIdQuery.ToString("X").PadLeft(4, '0'));
                }
            }

            UInt16 protocolIdentifier = (UInt16)(buffer[2] << 8 + buffer[3]);
            if (protocolIdentifier != 0)
            {
                Console.WriteLine("Invalid Protocol Identifier: {0:X}, expected: 0x0000", protocolIdentifier);
                throw new ModbusException("ModbusProtocolError: Invalid Protocol Identifier: 0x" + protocolIdentifier.ToString("X").PadLeft(4, '0') + ", expected: 0x0000");
            }

            UInt16 packetLen = (UInt16)((buffer[4] << 8) + buffer[5]);
            if (packetLen + 6 != length)
            {
                Console.WriteLine("Invalid Packet Length: {0}+6 != received: {1}", packetLen, length);
                throw new ModbusException("ModbusProtocolError: Invalid Packet Length: " + packetLen.ToString() + " + 6 (header) != received: " + length.ToString());
            }

            return true;
        }

        public byte[] readSerialCustom(int expected_len, int timeout)
        {
            // Custom function to force reading an expected number of bytes
            // or timeout elapse
            //
            // Ho visto che con alcuni convertitori USB/2323/485 in alcuni casi serial.readTimeout per
            // qualche motivo sembra non essere gestito e scadere prima del valore che vado a impostare
            // per cui gestisco anche monte il timeout

            byte[] output = new byte[256];

            int Length = 0;
            //byte[] output = new byte[256];

            serialPort.ReadTimeout = timeout;

            try
            {
                lock (lockModbus)
                {
                    byte[] buffer = new byte[256];
                    int len_tmp = 0;

                    DateTimeOffset start = DateTimeOffset.UtcNow;
                    long epoch = start.ToUnixTimeMilliseconds();

                    while (Length < expected_len)
                    {
                        // Normal read
                        try
                        {
                            len_tmp = serialPort.Read(buffer, 0, expected_len); // len_tmp sometimes could be < expected_len
                        }
                        catch
                        {
                            len_tmp = 0;
                        }

                        if (len_tmp > 0)
                        {
                            Array.Copy(buffer, 0, output, Length, len_tmp);
                            Length += len_tmp;
                        }


                        // Timeout elapsed
                        DateTimeOffset now = DateTimeOffset.UtcNow;
                        long epochNow = now.ToUnixTimeMilliseconds();

                        if ((epochNow - epoch) > timeout)
                        {
                            return new byte[0];
                        }
                    }

                    Array.Resize(ref output, Length);
                }
                return output;
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
                return new byte[0];
            }
        }

        public UInt16[] readCoilStatus_01(byte slave_add, uint start_add, uint no_of_coils, int readTimeout)
        {
            byte[] query;
            byte[] response;
            UInt16[] result = new UInt16[no_of_coils];

            Count_FC01++;

            if ((type == def.TYPE_TCP_SOCK || type == def.TYPE_TCP_REOPEN || type == def.TYPE_TCP_SECURE) && ClientActive)
            {
                /*
               TCP:
               0x00
               0x01
               0x00
               0x00
               0x00 -> Message Length Hi
               0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

               0x07 -> Slave Address
               0x01 -> Function
               0x01 -> Start Addr Hi
               0x2C -> Start Addr Lo
               0x00 -> No of Registers Hi
               0x03 -> No of Registers Lo
                */

                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x01;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);
                query[10] = (byte)(no_of_coils >> 8);
                query[11] = (byte)(no_of_coils);

                if (type == def.TYPE_TCP_REOPEN)
                {
                    client = new TcpClient(ip_address, int.Parse(port));
                    stream = client.GetStream();
                }

                TX_set = true;
                int Length = 0;

                if (type == def.TYPE_TCP_SECURE)
                {
                    lock (lockModbus)
                    {
                        sslStream.WriteTimeout = readTimeout;
                        sslStream.ReadTimeout = readTimeout;

                        response = new Byte[buffer_dimension];

                        // Flush di eventuali dati presenti
                        sslStream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = sslStream.Read(response, 0, response.Length);
                    }
                }
                else
                {
                    lock (lockModbus)
                    {
                        stream.WriteTimeout = readTimeout;
                        stream.ReadTimeout = readTimeout;

                        response = new Byte[buffer_dimension];

                        // Flush di eventuali dati presenti
                        if (type == def.TYPE_TCP_SOCK && stream.DataAvailable)
                            stream.Read(response, 0, response.Length);

                        stream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = stream.Read(response, 0, response.Length);
                    }
                }

                if(type == def.TYPE_TCP_REOPEN)
                    client.Close();

                // Timeout
                if (Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                // CheckModbusTcp Answer
                if (!checkModbusTcpResponse(response, Length, query))
                {
                    Console_print(" CheckModbusResponse failed", null, 0);
                }

                RX_set = true;

                Console_printByte("Rx: ", response, Length);
                Console_print(" Rx <- ", response, Length);

                // Modbus Error Code
                if ((response[7] & 0x80) > 0)
                {
                    int errCode = response[8];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                // Leggo i bit di ciascun byte partendo dal 9 che contiene le prime 8 coils
                // La coil 0 e nel LSb, la coil 7 nel MSb del primo byte, la 8 nel LSb del secondo byte
                for (int i = 9; i < Length; i += 1)
                {
                    for (int a = 0; a < 8; a++)
                    {
                        if ((i - 9) * 8 + a >= no_of_coils)
                            break;

                        result[(i - 9) * 8 + a] = (response[i] & (1 << a)) > 0 ? (byte)(1) : (byte)(0);
                    }
                }

                // debug
                // Console.WriteLine("Result (array of coils): " + result);

                return result;

            }
            else if (type == def.TYPE_RTU && ClientActive)
            {
                /*
                RTU
                0x07 -> Slave Address
                0x01 -> Function
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x01;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);
                query[4] = (byte)(no_of_coils >> 8);
                query[5] = (byte)(no_of_coils);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                TX_set = true;       // pictureBox gialla

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Console_printByte("Tx: ", query, query.Length);
                Console_print(" Tx -> ", query, query.Length);

                response = new Byte[buffer_dimension];

                try
                {
                    response = readSerialCustom((UInt16)((no_of_coils / 8)) + (no_of_coils % 8 > 0 ? 1 : 0) + 5, readTimeout);
                }
                catch
                {
                    response = new byte[]{ };
                }

                // Timeout
                if (response.Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                RX_set = true;        // pictureBox gialla

                Console_printByte("Rx: ", response, response.Length);
                Console_print(" Rx <- ", response, response.Length);

                // Check CRC
                if (!Check_CRC(response, response.Length))
                {
                    throw new ModbusException("CRC Error");
                }

                // Modbus Error Code
                if ((response[1] & 0x80) > 0)
                {
                    int errCode = response[2];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                // Leggo i bit di ciascun byte partendo dal 3 che contiene le prime 8 coils
                // La coil 0 e' nel LSb, la coil 7 nel MSb del primo byte, la 8 nel LSb del secondo byte
                for (int i = 3; i < (response.Length - 2); i++)
                {
                    for (int a = 0; a < 8; a++)
                    {
                        if ((i - 3) * 8 + a >= no_of_coils)
                            break;

                        result[(i - 3) * 8 + a] = (response[i] & (1 << a)) > 0 ? (byte)(1) : (byte)(0);
                    }
                }

                return result;

            }
            else
            {
                Console.WriteLine("No Modbus context available");
                return null;
            }
        }

        public UInt16[] readInputStatus_02(byte slave_add, uint start_add, uint no_of_input, int readTimeout)
        {
            byte[] query;
            byte[] response;
            UInt16[] result = new UInt16[no_of_input];

            Count_FC02++;

            if ((type == def.TYPE_TCP_SOCK || type == def.TYPE_TCP_REOPEN || type == def.TYPE_TCP_SECURE) && ClientActive)
            {
                /*
                TCP:
                0x00
                0x01
                0x00
                0x00
                0x00 -> Message Length Hi
                0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

                0x07 -> Slave Address
                0x02 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                */

                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x02;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);
                query[10] = (byte)(no_of_input >> 8);
                query[11] = (byte)(no_of_input);

                if (type == def.TYPE_TCP_REOPEN)
                {
                    client = new TcpClient(ip_address, int.Parse(port));
                    stream = client.GetStream();
                }

                TX_set = true;

                response = new Byte[buffer_dimension];
                int Length = 0;

                if (type == def.TYPE_TCP_SECURE)
                {
                    lock (lockModbus)
                    {
                        sslStream.WriteTimeout = readTimeout;
                        sslStream.ReadTimeout = readTimeout;

                        sslStream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = sslStream.Read(response, 0, response.Length);
                    }
                }
                else
                {
                    lock (lockModbus)
                    {
                        stream.WriteTimeout = readTimeout;
                        stream.ReadTimeout = readTimeout;

                        // Flush di eventuali dati presenti
                        if (type == def.TYPE_TCP_SOCK && stream.DataAvailable)
                            stream.Read(response, 0, response.Length);

                        stream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = stream.Read(response, 0, response.Length);
                    }
                }

                if (type == def.TYPE_TCP_REOPEN)
                    client.Close();

                // Timeout
                if (Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                // CheckModbusTcp Answer
                if (!checkModbusTcpResponse(response, Length, query))
                {
                    Console_print(" CheckModbusResponse failed", null, 0);
                }

                RX_set = true;

                Console_printByte("Rx: ", response, Length);
                Console_print(" Rx <- ", response, Length);

                // Modbus Error Code
                if ((response[7] & 0x80) > 0)
                {
                    int errCode = response[8];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                // Leggo i bit di ciascun byte partendo dal 9 che contiene le prime 8 coils
                // La coil 0 e nel LSb, la coil 7 nel MSb del primo byte, la 8 nel LSb del secondo byte
                for (int i = 9; i < Length; i += 1)
                {
                    for (int a = 0; a < 8; a++)
                    {
                        if ((i - 9) * 8 + a >= no_of_input)
                            break;

                        result[(i - 9) * 8 + a] = (response[i] & (1 << a)) > 0 ? (byte)(1) : (byte)(0);
                    }
                }

                // debug
                // Console.WriteLine("Result (array of inputs): " + result);

                return result;

            }
            else if (type == def.TYPE_RTU && ClientActive)
            {
                /*
                RTU
                0x07 -> Slave Address
                0x02 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x02;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);
                query[4] = (byte)(no_of_input >> 8);
                query[5] = (byte)(no_of_input);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                TX_set = true;       // pictureBox gialla

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Console_printByte("Tx: ", query, query.Length);
                Console_print(" Tx -> ", query, query.Length);

                response = new Byte[buffer_dimension];

                try
                {
                    response = readSerialCustom((UInt16)((no_of_input / 8)) + (no_of_input % 8 > 0 ? 1 : 0) + 5, readTimeout);
                }
                catch
                {
                    // debug
                    // Console.WriteLine("Result (array of inputs): " + result);
                }

                if (response.Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                RX_set = true;        // pictureBox gialla

                Console_printByte("Rx: ", response, response.Length);
                Console_print(" Rx <- ", response, response.Length);

                // Check CRC
                if (!Check_CRC(response, response.Length))
                {
                    throw new ModbusException("CRC Error");
                }

                // Modbus Error Code
                if ((response[1] & 0x80) > 0)
                {
                    int errCode = response[2];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                // Leggo i bit di ciascun byte partendo dal 9 che contiene le prime 8 coils
                // La coil 0 e nel LSb, la coil 7 nel MSb del primo byte, la 8 nel LSb del secondo byte
                for (int i = 3; i < (response.Length - 2); i++)
                {
                    for (int a = 0; a < 8; a++)
                    {
                        if ((i - 3) * 8 + a >= no_of_input)
                            break;

                        result[(i - 3) * 8 + a] = (response[i] & (1 << a)) > 0 ? (byte)(1) : (byte)(0);
                    }
                }

                // debug
                // Console.WriteLine("Result (array of inputs): " + result);

                return result;
            }
            else
            {
                Console.WriteLine("No Modbus context available");
                return null;
            }
        }

        public UInt16[] readHoldingRegister_03(byte slave_add, uint start_add, uint no_of_registers, int readTimeout)
        {
            byte[] query;
            byte[] response;
            UInt16[] result = new UInt16[no_of_registers];

            Count_FC03++;

            if ((type == def.TYPE_TCP_SOCK || type == def.TYPE_TCP_REOPEN || type == def.TYPE_TCP_SECURE) && ClientActive)
            {
                /*
                TCP:
                0x00
                0x01
                0x00
                0x00
                0x00 -> Message Length Hi
                0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

                0x07 -> Slave Address
                0x03 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                 */

                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x03;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);
                query[10] = (byte)(no_of_registers >> 8);
                query[11] = (byte)(no_of_registers);

                if (type == def.TYPE_TCP_REOPEN)
                {
                    client = new TcpClient(ip_address, int.Parse(port));
                    stream = client.GetStream();
                }

                TX_set = true;       // pictureBox gialla

                response = new Byte[buffer_dimension];
                int Length = 0;

                if (type == def.TYPE_TCP_SECURE)
                {
                    lock (lockModbus)
                    {
                        sslStream.WriteTimeout = readTimeout;
                        sslStream.ReadTimeout = readTimeout;

                        sslStream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = sslStream.Read(response, 0, response.Length);
                    }
                }
                else
                {
                    lock (lockModbus)
                    {
                        stream.WriteTimeout = readTimeout;
                        stream.ReadTimeout = readTimeout;

                        // Flush di eventuali dati presenti
                        if (type == def.TYPE_TCP_SOCK && stream.DataAvailable)
                            stream.Read(response, 0, response.Length);

                        stream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = stream.Read(response, 0, response.Length);
                    }
                }

                if (type == def.TYPE_TCP_REOPEN)
                    client.Close();

                // Timeout
                if (Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                // CheckModbusTcp Answer
                if (!checkModbusTcpResponse(response, Length, query))
                {
                    Console_print(" CheckModbusResponse failed", null, 0);
                }

                RX_set = true;       // pictureBox gialla

                Console_printByte("Rx: ", response, Length);
                Console_print(" Rx <- ", response, Length);

                // Modbus Error Code
                if ((response[7] & 0x80) > 0)
                {
                    int errCode = response[8];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                for (int i = 9; i < Length; i += 2)
                {
                    result[(i - 9) / 2] = (UInt16)((UInt16)(response[i] << 8) + (UInt16)(response[i + 1]));
                }

                // debug
                //Console.WriteLine("Result (array of registers): " + result);

                return result;

            }
            else if (type == def.TYPE_RTU && ClientActive)
            {
                /*
                RTU
                0x07 -> Slave Address
                0x03 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x03;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);
                query[4] = (byte)(no_of_registers >> 8);
                query[5] = (byte)(no_of_registers);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                TX_set = true;       // pictureBox gialla

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Console_printByte("Tx: ", query, query.Length);
                Console_print(" Tx -> ", query, query.Length);

                response = new Byte[buffer_dimension];

                try
                {
                    response = readSerialCustom((int)(no_of_registers * 2) + 5, readTimeout);
                }
                catch
                {
                    response = new byte[] { };
                }

                // Timeout
                if (response.Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                RX_set = true;        // pictureBox gialla

                Console_printByte("Rx: ", response, response.Length);
                Console_print(" Rx <- ", response, response.Length);

                for (int i = 3; i < response.Length - 2; i += 2)
                {
                    result[(i - 3) / 2] = (UInt16)((UInt16)(response[i] << 8) + (UInt16)(response[i + 1]));
                }

                // Check CRC
                if (!Check_CRC(response, response.Length))
                {
                    throw new ModbusException("CRC Error");
                }

                // Modbus Error Code
                if ((response[1] & 0x80) > 0)
                {
                    int errCode = response[2];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                return result;
            }
            else
            {
                string[] error = { "?" };
                Console.WriteLine("No Modbus context available");
                return null;
            }
        }

        public UInt16[] readInputRegister_04(byte slave_add, uint start_add, uint no_of_registers, int readTimeout)
        {

            byte[] query;
            byte[] response;
            UInt16[] result = new UInt16[no_of_registers];

            Count_FC04++;

            if ((type == def.TYPE_TCP_SOCK || type == def.TYPE_TCP_REOPEN || type == def.TYPE_TCP_SECURE) && ClientActive)
            {
                /*
                TCP:
                0x00
                0x01
                0x00
                0x00
                0x00 -> Message Length Hi
                0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

                0x07 -> Slave Address
                0x04 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                */

                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x04;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);
                query[10] = (byte)(no_of_registers >> 8);
                query[11] = (byte)(no_of_registers);

                if (type == def.TYPE_TCP_REOPEN)
                {
                    client = new TcpClient(ip_address, int.Parse(port));
                    stream = client.GetStream();
                }

                TX_set = true;       // pictureBox gialla

                response = new Byte[buffer_dimension];
                int Length = 0;

                if (type == def.TYPE_TCP_SECURE)
                {
                    lock (lockModbus)
                    {
                        sslStream.WriteTimeout = readTimeout;
                        sslStream.ReadTimeout = readTimeout;

                        sslStream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = sslStream.Read(response, 0, response.Length);
                    }
                }
                else
                {
                    lock (lockModbus)
                    {
                        stream.WriteTimeout = readTimeout;
                        stream.ReadTimeout = readTimeout;

                        // Flush di eventuali dati presenti
                        if (type == def.TYPE_TCP_SOCK && stream.DataAvailable)
                            stream.Read(response, 0, response.Length);

                        stream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = stream.Read(response, 0, response.Length);
                    }
                }

                if (type == def.TYPE_TCP_REOPEN)
                    client.Close();

                // Timeout
                if (Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                // CheckModbusTcp Answer
                if (!checkModbusTcpResponse(response, Length, query))
                {
                    Console_print(" CheckModbusResponse failed", null, 0);
                }

                RX_set = true;       // pictureBox gialla

                Console_printByte("Rx: ", response, Length);
                Console_print(" Rx <- ", response, Length);

                // Modbus Error Code
                if ((response[7] & 0x80) > 0)
                {
                    int errCode = response[8];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                for (int i = 9; i < Length; i += 2)
                {
                    result[(i - 9) / 2] = (UInt16)((UInt16)(response[i] << 8) + (UInt16)(response[i + 1]));
                }

                // debug
                //Console.WriteLine("Result (array of registers): " + result);

                return result;

            }
            else if (type == def.TYPE_RTU && ClientActive)
            {
                /*
                RTU:
                0x07 -> Slave Address
                0x04 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x04;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);
                query[4] = (byte)(no_of_registers >> 8);
                query[5] = (byte)(no_of_registers);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                TX_set = true;       // pictureBox gialla

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Thread.Sleep(200);

                Console_printByte("Tx: ", query, query.Length);
                Console_print(" Tx -> ", query, query.Length);

                response = new Byte[buffer_dimension];

                try
                {
                    response = readSerialCustom((UInt16)(no_of_registers * 2) + 5, readTimeout);
                }
                catch
                {
                    response = new byte[] { };
                }

                // Timeout
                if (response.Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                RX_set = true;        // pictureBox gialla

                Console_printByte("Rx: ", response, response.Length);
                Console_print(" Rx <- ", response, response.Length);

                for (int i = 3; i < response.Length - 2; i += 2) //-2 di CRC
                {
                    result[(i - 3) / 2] = (UInt16)((UInt16)(response[i] << 8) + (UInt16)(response[i + 1]));
                }

                // Check CRC
                if (!Check_CRC(response, response.Length))
                {
                    throw new ModbusException("CRC Error");
                }

                // Modbus Error Code
                if ((response[1] & 0x80) > 0)
                {
                    int errCode = response[2];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                return result;
            }
            else
            {
                Console.WriteLine("No Modbus context available");
                return null;
            }
        }

        public bool? forceSingleCoil_05(byte slave_add, uint start_add, uint state, int readTimeout)
        {
            //True se la funzione riceve risposta affermativa
            
            byte[] query;
            byte[] response;
            //uint[] result = new uint[no_of_registers];

            Count_FC05++;

            if ((type == def.TYPE_TCP_SOCK || type == def.TYPE_TCP_REOPEN || type == def.TYPE_TCP_SECURE) && ClientActive)
            {
                /*
                TCP:
                0x00
                0x01
                0x00
                0x00
                0x00 -> Message Length Hi
                0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

                0x07 -> Slave Address
                0x05 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi (0xFF -> On, 0x00 -> Off)
                0x03 -> No of Registers Lo (Sempre 0x00)
                 */

                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;    //Message Length
                query[5] = 0x06;    //Message Length

                query[6] = slave_add;
                query[7] = 0x05;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);

                if (state > 0)
                    query[10] = 0xFF;
                else
                    query[10] = 0x00;

                query[11] = 0x00;

                if (type == def.TYPE_TCP_REOPEN)
                {
                    client = new TcpClient(ip_address, int.Parse(port));
                    stream = client.GetStream();
                }

                TX_set = true;       // pictureBox gialla

                response = new Byte[buffer_dimension];
                int Length = 0;

                if (type == def.TYPE_TCP_SECURE)
                {
                    lock (lockModbus)
                    {
                        sslStream.WriteTimeout = readTimeout;
                        sslStream.ReadTimeout = readTimeout;

                        sslStream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = sslStream.Read(response, 0, response.Length);
                    }
                }
                else
                {
                    lock (lockModbus)
                    {
                        stream.WriteTimeout = readTimeout;
                        stream.ReadTimeout = readTimeout;

                        // Flush di eventuali dati presenti
                        if (type == def.TYPE_TCP_SOCK && stream.DataAvailable)
                            stream.Read(response, 0, response.Length);

                        stream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = stream.Read(response, 0, response.Length);
                    }
                }

                if (type == def.TYPE_TCP_REOPEN)
                    client.Close();

                // Timeout
                if (Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                // CheckModbusTcp Answer
                if (!checkModbusTcpResponse(response, Length, query))
                {
                    Console_print(" CheckModbusResponse failed", null, 0);
                }

                RX_set = true;        // pictureBox gialla

                Console_printByte("Rx: ", response, Length);
                Console_print(" Rx <- ", response, Length);

                // Modbus Error Code
                if ((response[7] & 0x80) > 0)
                {
                    int errCode = response[8];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                // Check response
                if (response[6] == query[6] &&      // Slave ID
                    response[7] == query[7] &&      // FC
                    response[8] == query[8] &&      // Start Addr
                    response[9] == query[9] &&      // Start Addr
                    response[10] == query[10] &&    // Value
                    response[11] == query[11])      // Value
                {
                    return true;
                }
                    
                return false;

            }
            else if (type == def.TYPE_RTU && ClientActive)
            {
                //True se la funzione riceve risposta affermativa

                /*
                RTU
                0x07 -> Slave Address
                0x05 -> Header
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi (0xFF -> On, 0x00 -> Off)
                0x03 -> No of Registers Lo (Sempre 0x00)
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                 */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x05;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);

                if (state > 0)
                    query[4] = 0xFF;
                else
                    query[4] = 0x00;

                query[5] = 0x00;

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                TX_set = true;       // pictureBox gialla

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Console_printByte("Tx: ", query, query.Length);
                Console_print(" Tx -> ", query, query.Length);

                try
                {
                    response = readSerialCustom(8, readTimeout);
                }
                catch
                {
                    response = new byte[] { };
                }

                // Timeout
                if (response.Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                RX_set = true;        // pictureBox gialla

                Console_printByte("Rx: ", response, response.Length);
                Console_print(" Rx <- ", response, response.Length);

                // Check CRC
                if (!Check_CRC(response, response.Length))
                {
                    throw new ModbusException("CRC Error");
                }

                // Modbus Error Code
                if ((response[1] & 0x80) > 0)
                {
                    int errCode = response[2];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                // Check response
                if (response[0] == query[0] &&    // Slave ID
                    response[1] == query[1] &&    // FC
                    response[2] == query[2] &&    // Start Addr
                    response[3] == query[3] &&    // Start Addr
                    response[4] == query[4] &&    // Value
                    response[5] == query[5])      // Value
                {
                    return true;
                }

                return false;
            }
            else
            {
                Console.WriteLine("No Modbus context available");
                return false;
            }
        }

        public bool? presetSingleRegister_06(byte slave_add, uint start_add, uint value, int readTimeout)
        {
            //True se la funzione riceve risposta affermativa

            byte[] query;
            byte[] response;

            Count_FC06++;

            if ((type == def.TYPE_TCP_SOCK || type == def.TYPE_TCP_REOPEN || type == def.TYPE_TCP_SECURE) && ClientActive)
            {
                /*
                TCP:
                0x00
                0x01
                0x00
                0x00
                0x00 -> Message Length Hi
                0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

                0x07 -> Slave Address
                0x06 -> FUnction Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                 */

                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x06;
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);

                query[10] = (byte)(value >> 8);
                query[11] = (byte)(value);

                if (type == def.TYPE_TCP_REOPEN)
                {
                    client = new TcpClient(ip_address, int.Parse(port));
                    stream = client.GetStream();
                }

                TX_set = true;       // pictureBox gialla

                response = new Byte[buffer_dimension];
                int Length = 0;

                if (type == def.TYPE_TCP_SECURE)
                {
                    lock (lockModbus)
                    {
                        sslStream.WriteTimeout = readTimeout;
                        sslStream.ReadTimeout = readTimeout;

                        sslStream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = sslStream.Read(response, 0, response.Length);
                    }
                }
                else
                {
                    lock (lockModbus)
                    {
                        stream.WriteTimeout = readTimeout;
                        stream.ReadTimeout = readTimeout;

                        // Flush di eventuali dati presenti
                        if (type == def.TYPE_TCP_SOCK && stream.DataAvailable)
                            stream.Read(response, 0, response.Length);

                        stream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = stream.Read(response, 0, response.Length);
                    }
                }

                if (type == def.TYPE_TCP_REOPEN)
                    client.Close();

                // Timeout
                if (Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                // CheckModbusTcp Answer
                if (!checkModbusTcpResponse(response, Length, query))
                {
                    Console_print(" CheckModbusResponse failed", null, 0);
                }

                RX_set = true;        // pictureBox gialla

                Console_printByte("Rx: ", response, Length);
                Console_print(" Rx <- ", response, Length);

                // Modbus Error Code
                if ((response[7] & 0x80) > 0)
                {
                    int errCode = response[8];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                // Check response
                if (response[6] == query[6] &&      // Slave ID
                    response[7] == query[7] &&      // FC
                    response[8] == query[8] &&      // Start Addr
                    response[9] == query[9] &&      // Start Addr
                    response[10] == query[10] &&    // Value
                    response[11] == query[11])      // Value
                {
                    return true;
                }

                return false;

            }
            else if (type == def.TYPE_RTU && ClientActive)
            {
                //True se la funzione riceve risposta affermativa

                /*
                RTU
                0x07 -> Slave Address
                0x06 -> Header
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x06;
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);

                query[4] = (byte)(value >> 8);
                query[5] = (byte)(value);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                TX_set = true;       // pictureBox gialla

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Console_printByte("Tx: ", query, query.Length);
                Console_print(" Tx -> ", query, query.Length);

                try
                {
                    response = readSerialCustom(8, readTimeout);
                }
                catch
                {
                    response = new byte[] { };
                }

                if (response.Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                RX_set = true;        // pictureBox gialla

                Console_printByte("Rx: ", response, response.Length);
                Console_print(" Rx <- ", response, response.Length);

                // Check CRC
                if (!Check_CRC(response, response.Length))
                {
                    throw new ModbusException("CRC Error");
                }

                // Modbus Error Code
                if ((response[1] & 0x80) > 0)
                {
                    int errCode = response[2];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                // Check response
                if (response[0] == query[0] &&    // Slave ID
                    response[1] == query[1] &&    // FC
                    response[2] == query[2] &&    // Start Addr
                    response[3] == query[3] &&    // Start Addr
                    response[4] == query[4] &&    // Value
                    response[5] == query[5])      // Value
                {
                    return true;
                }

                return false;
            }
            else
            {
                Console.WriteLine("No Modbus context available");
                return false;
            }
        }

        public String diagnostics_08(byte slave_add, uint sub_func, uint data, int readTimeout)
        {

            byte[] query;
            byte[] response;
            String[] result = new String[buffer_dimension];

            Count_FC08++;

            if ((type == def.TYPE_TCP_SOCK || type == def.TYPE_TCP_REOPEN || type == def.TYPE_TCP_SECURE) && ClientActive)
            {
                /*
                TCP:
                0x00
                0x01
                0x00
                0x00
                0x00 -> Message Length Hi
                0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

                0x07 -> Slave Address
                0x04 -> Function Code
                0x01 -> Subfunction Hi
                0x2C -> Subfunction LO
                0x00 -> Data Hi
                0x03 -> Data Lo
                */

                queryCounter++;
                query = new byte[12];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                query[4] = 0x00;
                query[5] = 0x06;

                query[6] = slave_add;
                query[7] = 0x08;
                query[8] = (byte)(sub_func >> 8);
                query[9] = (byte)(sub_func);
                query[10] = (byte)(data >> 8);
                query[11] = (byte)(data);

                if (type == def.TYPE_TCP_REOPEN)
                {
                    client = new TcpClient(ip_address, int.Parse(port));
                    stream = client.GetStream();
                }

                TX_set = true;       // pictureBox gialla

                response = new Byte[buffer_dimension];
                int Length = 0;

                if (type == def.TYPE_TCP_SECURE)
                {
                    lock (lockModbus)
                    {
                        sslStream.WriteTimeout = readTimeout;
                        sslStream.ReadTimeout = readTimeout;

                        sslStream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = sslStream.Read(response, 0, response.Length);
                    }
                }
                else
                {
                    lock (lockModbus)
                    {
                        stream.WriteTimeout = readTimeout;
                        stream.ReadTimeout = readTimeout;

                        // Flush di eventuali dati presenti
                        if (type == def.TYPE_TCP_SOCK && stream.DataAvailable)
                            stream.Read(response, 0, response.Length);

                        stream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = stream.Read(response, 0, response.Length);
                    }
                }

                if (type == def.TYPE_TCP_REOPEN)
                    client.Close();

                // Timeout
                if (Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                // CheckModbusTcp Answer
                if (!checkModbusTcpResponse(response, Length, query))
                {
                    Console_print(" CheckModbusResponse failed", null, 0);
                }

                RX_set = true;       // pictureBox gialla

                Console_printByte("Rx: ", response, Length);
                Console_print(" Rx <- ", response, Length);

                // Modbus Error Code
                if ((response[7] & 0x80) > 0)
                {
                    int errCode = response[8];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                // debug
                //Console.WriteLine("Result (array of registers): " + result);

                string result_ = "";
                result = new String[Length];

                for (int i = 0; i < Length; i++)
                {
                    try
                    {
                        //if(result[i] < 10)
                        result[i] = response[i].ToString("X");

                        if (int.Parse(result[i]) < 10)
                            result_ += "0";

                        result_ += response[i].ToString("X") + " ";
                    }
                    catch
                    {
                        Console.WriteLine("Errore analisi response diagnostic function");
                    }
                }

                return result_;

            }
            else if (type == def.TYPE_RTU && ClientActive)
            {
                /*
                RTU:
                0x07 -> Slave Address
                0x04 -> Function Code
                0x01 -> Subfunction Hi
                0x2C -> Subfunction Lo
                0x00 -> Data Hi
                0x03 -> Data Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[8];

                query[0] = slave_add;
                query[1] = 0x08;
                query[2] = (byte)(sub_func >> 8);
                query[3] = (byte)(sub_func);
                query[4] = (byte)(data >> 8);
                query[5] = (byte)(data);

                byte[] crc = Calcolo_CRC(query, 6);

                query[6] = crc[0];
                query[7] = crc[1];

                TX_set = true;       // pictureBox gialla

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Console_printByte("Tx: ", query, query.Length);
                Console_print(" Tx -> ", query, query.Length);

                int Length = 0;
                response = new Byte[buffer_dimension];

                try
                {
                    Length = serialPort.Read(response, 0, response.Length);
                }
                catch
                {
                    Console.WriteLine("Timeout lettura porta seriale");
                }

                // Check CRC
                if (!Check_CRC(response, Length))
                {
                    throw new ModbusException("CRC Error");
                }

                // Timeout
                if (Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                RX_set = true;       // pictureBox gialla

                Console_printByte("Rx: ", response, Length);
                Console_print(" Rx <- ", response, Length);

                // Modbus Error Code
                if ((response[1] & 0x80) > 0)
                {
                    int errCode = response[2];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                string result_ = "";
                result = new String[Length];

                for (int i = 0; i < Length; i++)
                {
                    try
                    { 
                        result[i] = response[i].ToString("X");

                        if (int.Parse(result[i]) < 10)
                            result_ += "0";

                        result_ += response[i].ToString("X") + " ";
                    }
                    catch
                    {
                        Console.WriteLine("Errore analisi response diagnostic function");
                    }
                }

                return result_;
            }
            else
            {
                Console.WriteLine("No Modbus context available");
                return null;
            }
        }

        public bool? forceMultipleCoils_15(byte slave_add, uint start_add, bool[] coils_value, int readTimeout)
        {
            // True se la funzione riceve risposta affermativa

            byte[] query;
            byte[] response;

            Count_FC15++;

            if ((type == def.TYPE_TCP_SOCK || type == def.TYPE_TCP_REOPEN || type == def.TYPE_TCP_SECURE) && ClientActive)
            {
                /*
               TCP:
               0x00
               0x01
               0x00
               0x00
               0x00 -> Message Length Hi
               0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

               0x07 -> Slave Address
               0x15 -> Function Code
               0x01 -> Start Addr Hi
               0x2C -> Start Addr Lo
               0x00 -> No of Registers Hi
               0x03 -> No of Registers Lo

               0x02 -> Byte count

               0x00 -> Data Hi
               0x00 -> Data Lo
                */

                queryCounter++;
                query = new byte[13 + (coils_value.Length / 8) + (coils_value.Length % 8 == 0 ? 0 : 1)];

                //Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                //Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                //Bytes to follow
                query[4] = 0x00;
                query[5] = (byte)(0x07 + (coils_value.Length / 8) + (coils_value.Length % 8 == 0 ? 0x00 : 0x01));

                query[6] = slave_add;
                query[7] = 0x0F;            // FC Code

                // Starting address
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);

                // Number of regsiters
                query[10] = (byte)(coils_value.Length >> 8);
                query[11] = (byte)(coils_value.Length);

                // Byte count
                query[12] = (byte)((coils_value.Length / 8) + (coils_value.Length % 8 == 0 ? 0 : 1));

                for (int i = 0; i < ((coils_value.Length / 8) + (coils_value.Length % 8 == 0 ? 0 : 1)); i++)
                {
                    byte val = 0;

                    for(int a = 0; a < 8; a++)
                    {
                        if (a + i * 8 < coils_value.Length)
                        {
                            if(coils_value[a + (i * 8)])
                            {
                                val += (byte)(1 << a);

                                // debug
                                Console.WriteLine("coil " + a.ToString() + ":1");
                            }
                            else
                            {
                                // debug
                                Console.WriteLine("coil " + a.ToString() + ":0");
                            }
                        }
                    }

                    query[13 + i] = val;

                    // debug
                    Console.WriteLine("byte: " +  val.ToString());
                }

                if (type == def.TYPE_TCP_REOPEN)
                {
                    client = new TcpClient(ip_address, int.Parse(port));
                    stream = client.GetStream();
                }

                TX_set = true;       // pictureBox gialla

                response = new Byte[buffer_dimension];
                int Length = 0;

                if (type == def.TYPE_TCP_SECURE)
                {
                    lock (lockModbus)
                    {
                        sslStream.WriteTimeout = readTimeout;
                        sslStream.ReadTimeout = readTimeout;

                        sslStream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = sslStream.Read(response, 0, response.Length);
                    }
                }
                else
                {
                    lock (lockModbus)
                    {
                        stream.WriteTimeout = readTimeout;
                        stream.ReadTimeout = readTimeout;

                        // Flush di eventuali dati presenti
                        if (type == def.TYPE_TCP_SOCK && stream.DataAvailable)
                            stream.Read(response, 0, response.Length);

                        stream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = stream.Read(response, 0, response.Length);
                    }
                }

                if (type == def.TYPE_TCP_REOPEN)
                    client.Close();

                // Timeout
                if (Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                // CheckModbusTcp Answer
                if (!checkModbusTcpResponse(response, Length, query))
                {
                    Console_print(" CheckModbusResponse failed", null, 0);
                }

                RX_set = true;        // pictureBox gialla

                Console_printByte("Rx: ", response, Length);
                Console_print(" Rx <- ", response, Length);

                // Modbus Error Code
                if ((response[7] & 0x80) > 0)
                {
                    int errCode = response[8];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                if (response.Length > 11)
                {
                    if (response[6] == query[6] &&      // Slave ID
                        response[7] == query[7] &&      // FC
                        response[8] == query[8] &&      // Start Addr
                        response[9] == query[9] &&      // Start Addr
                        response[10] == query[10] &&    // Nr of coils
                        response[11] == query[11])      // Nr of coils
                    {
                        return true;
                    }
                }

                return false;

            }
            else if (type == def.TYPE_RTU && ClientActive)
            {
                //True se la funzione riceve risposta affermativa

                /*
                RTU
                0x07 -> Slave Address
                0x06 -> FC
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x02 -> Byte count
                0x00 -> Data[0] Hi
                0x03 -> Data[0] Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                int limitCount = (coils_value.Length / 8) + (coils_value.Length % 8 == 0 ? 0 : 1);

                query = new byte[7 + limitCount + 2];

                query[0] = slave_add;
                query[1] = 0x0F;

                // Starting address
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);

                // Number of regsiters
                query[4] = (byte)(coils_value.Length >> 8);
                query[5] = (byte)(coils_value.Length);

                // Byte count
                query[6] = (byte)(limitCount);

                for (int i = 0; i < limitCount; i++)
                {
                    byte val = 0;

                    for (int a = 0; a < 8; a++)
                    {
                        if (a + i * 8 < coils_value.Length)
                        {
                            if (coils_value[a + i * 8])
                            {
                                val += (byte)(1 << a);

                                // debug
                                Console.WriteLine("coil " + a.ToString() + ":1");
                            }
                            else
                            {
                                // debug
                                Console.WriteLine("coil " + a.ToString() + ":0");
                            }
                        }
                    }

                    query[7 + i] = val;

                    // debug
                    Console.WriteLine("byte: " + val.ToString());
                }

                byte[] crc = Calcolo_CRC(query, 7 + limitCount);

                query[7 + limitCount] = crc[0];
                query[8 + limitCount] = crc[1];

                TX_set = true;       // pictureBox gialla

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Console_printByte("Tx: ", query, query.Length);
                Console_print(" Tx -> ", query, query.Length);

                response = new Byte[buffer_dimension];

                try
                {
                    response = readSerialCustom(8, readTimeout);
                }
                catch
                {
                    response = new byte[] { };
                }

                // Timeout
                if (response.Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                RX_set = true;        // pictureBox gialla

                Console_printByte("Rx: ", response, response.Length);
                Console_print(" Rx <- ", response, response.Length);

                // Check CRC
                if (!Check_CRC(response, response.Length))
                {
                    throw new ModbusException("CRC Error");
                }

                // Modbus Error Code
                if ((response[1] & 0x80) > 0)
                {
                    int errCode = response[2];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                // Check response
                if (response.Length > 5)
                {
                    if (response[0] == query[0] &&  // Slave ID
                        response[1] == query[1] &&  // FC
                        response[2] == query[2] &&  // Start Addr
                        response[3] == query[3] &&  // Start Addr
                        response[4] == query[4] &&  // Nr of coils
                        response[5] == query[5])    // Nr of coils
                    {
                        return true;
                    }
                }

                return false;
            }
            else
            {
                Console.WriteLine("No Modbus context available");
                return false;
            }
        }

        public UInt16[] presetMultipleRegisters_16(byte slave_add, uint start_add, UInt16[] register_value, int readTimeout)
        {
            // True se la funzione riceve risposta affermativa

            byte[] query;
            byte[] response;

            Count_FC16++;

            if ((type == def.TYPE_TCP_SOCK || type == def.TYPE_TCP_REOPEN || type == def.TYPE_TCP_SECURE) && ClientActive)
            {
                /*
                TCP:
                0x00
                0x01
                0x00
                0x00
                0x00 -> Message Length Hi
                0x06 -> Message Length Lo (Riferito ai 6 byte sottostanti)

                0x07 -> Slave Address
                0x16 -> Function Code
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
            
                0x02 -> Byte count

                0x00 -> Data Hi
                0x00 -> Data Lo
                 */

                queryCounter++;
                query = new byte[13 + register_value.Length*2];

                // Transaction identifier
                query[0] = (byte)(queryCounter >> 8);
                query[1] = (byte)(queryCounter);

                // Protocol identifier
                query[2] = 0x00;
                query[3] = 0x00;

                // Bytes to follow
                query[4] = (byte)(((register_value.Length * 2 + 7) << 8) & 0xFF);
                query[5] = (byte)((register_value.Length * 2 + 7) & 0xFF);

                query[6] = slave_add;
                query[7] = 0x10;

                // Starting address
                query[8] = (byte)(start_add >> 8);
                query[9] = (byte)(start_add);

                // Number of regsiters
                query[10] = (byte)(register_value.Length >> 8);
                query[11] = (byte)(register_value.Length);

                // Byte count
                query[12] = (byte)(register_value.Length * 2);

                for (int i = 0; i < register_value.Length; i++)
                {
                    query[13 + 2*i] = (byte)(register_value[i] >> 8);
                    query[14 + 2*i] = (byte)(register_value[i]);
                }

                if (type == def.TYPE_TCP_REOPEN)
                {
                    client = new TcpClient(ip_address, int.Parse(port));
                    stream = client.GetStream();
                }

                TX_set = true;       // pictureBox gialla

                response = new Byte[buffer_dimension];
                int Length = 0;

                if (type == def.TYPE_TCP_SECURE)
                {
                    lock (lockModbus)
                    {
                        sslStream.WriteTimeout = readTimeout;
                        sslStream.ReadTimeout = readTimeout;

                        sslStream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = sslStream.Read(response, 0, response.Length);
                    }
                }
                else
                {
                    lock (lockModbus)
                    {
                        stream.WriteTimeout = readTimeout;
                        stream.ReadTimeout = readTimeout;

                        // Flush di eventuali dati presenti
                        if (type == def.TYPE_TCP_SOCK && stream.DataAvailable)
                            stream.Read(response, 0, response.Length);

                        stream.Write(query, 0, query.Length);

                        Console_printByte("Tx: ", query, query.Length);
                        Console_print(" Tx -> ", query, query.Length);

                        Length = stream.Read(response, 0, response.Length);
                    }
                }

                if (type == def.TYPE_TCP_REOPEN)
                    client.Close();

                // Timeout
                if (Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                // CheckModbusTcp Answer
                if (!checkModbusTcpResponse(response, Length, query))
                {
                    Console_print(" CheckModbusResponse failed", null, 0);
                }

                RX_set = true;       // pictureBox gialla

                Console_printByte("Rx: ", response, Length);
                Console_print(" Rx <- ", response, Length);

                // Modbus Error Code
                if ((response[7] & 0x80) > 0)
                {
                    int errCode = response[8];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                if (response.Length > 11) {
                    if (response[6] == query[6] &&      // Slave ID
                        response[7] == query[7] &&      // FC
                        response[8] == query[8] &&      // Start Addr
                        response[9] == query[9] &&      // Start Addr
                        response[10] == query[10] &&    // Nr of regs
                        response[11] == query[11])      // Nr of regs
                    {
                        return register_value;
                    }
                }

                return new UInt16[0] { }; ;

            }
            else if (type == def.TYPE_RTU && ClientActive)
            {
                //True se la funzione riceve risposta affermativa

                /*
                RTU
                0x07 -> Slave Address
                0x06 -> Header
                0x01 -> Start Addr Hi
                0x2C -> Start Addr Lo
                0x00 -> No of Registers Hi
                0x03 -> No of Registers Lo
                0x02 -> Byte count
                0x00 -> Data[0] Hi
                0x03 -> Data[0] Lo
                0x?? -> CRC Hi
                0x?? -> CRC Lo
                */

                query = new byte[9 + register_value.Length*2];

                query[0] = slave_add;
                query[1] = 0x10;

                // Starting address
                query[2] = (byte)(start_add >> 8);
                query[3] = (byte)(start_add);

                // Number of regsiters
                query[4] = (byte)(register_value.Length >> 8);
                query[5] = (byte)(register_value.Length);

                // Byte count
                query[6] = (byte)(register_value.Length * 2);

                for (int i = 0; i < register_value.Length; i++)
                {
                    query[7 + 2 * i] = (byte)(register_value[i] >> 8);
                    query[8 + 2 * i] = (byte)(register_value[i]);
                }

                byte[] crc = Calcolo_CRC(query, 7 + register_value.Length*2);

                query[7 + register_value.Length * 2] = crc[0];
                query[8 + register_value.Length * 2] = crc[1];

                TX_set = true;       // pictureBox gialla

                serialPort.DiscardInBuffer();
                serialPort.DiscardOutBuffer();
                serialPort.ReadTimeout = readTimeout;
                serialPort.Write(query, 0, query.Length);

                Console_printByte("Tx: ", query, query.Length);
                Console_print(" Tx -> ", query, query.Length);

                try
                {
                    response = readSerialCustom(8, readTimeout); 
                }
                catch
                {
                    response = new byte[] { };
                }

                // Timeout
                if (response.Length == 0)
                {
                    Console_print(" Timed out", null, 0);
                    throw new ModbusException("Timed out");
                }

                RX_set = true;       // pictureBox gialla

                Console_printByte("Rx: ", response, response.Length);
                Console_print(" Rx <- ", response, response.Length);

                // Check CRC
                if (!Check_CRC(response, response.Length))
                {
                    throw new ModbusException("CRC Error");
                }

                // Modbus Error Code
                if ((response[1] & 0x80) > 0)
                {
                    int errCode = response[2];

                    Console_print(" ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode], null, 0);
                    throw new ModbusException("ModBus ErrCode: " + errCode.ToString() + " - " + ModbusErrorCodes[errCode]);
                }

                if (response.Length > 5)
                {
                    if (response[0] == query[0] &&  // Slave ID
                        response[1] == query[1] &&  // FC
                        response[2] == query[2] &&  // Start Addr
                        response[3] == query[3] &&  // Start Addr
                        response[4] == query[4] &&  // Nr of regs
                        response[5] == query[5])    // Nr of regs
                    {
                        return register_value;
                    }
                }

                return new UInt16[0] { };
            }
            else
            {
                Console.WriteLine("No Modbus context available");
                return new UInt16[0] { };
            }
        }

        //-----------------------------------------------------------------
        //--------------------Calcolo CRC 16 MODBUS------------------------
        //-----------------------------------------------------------------

        // Calcolo CRC MODBUS
        public byte[] Calcolo_CRC(byte[] message, int length)
        {
            UInt16 crc = 0xFFFF;
            byte[] result = new byte[2];

            for (int pos = 0; pos < length; pos++)
            {
                crc ^= (UInt16)message[pos];    //XOR

                for (int i = 8; i != 0; i--)
                {
                    // Passo ogni byte del pacchetto
                    if ((crc & 0x0001) != 0)
                    {
                        crc >>= 1;
                        crc ^= 0xA001;
                    }
                    else
                    {             
                        crc >>= 1;
                    }
                }
            }

            result[0] = (byte)(crc);        //LSB
            result[1] = (byte)(crc >> 8);   //MSB

            return result;
        }
        
        bool Check_CRC(byte[] message, int length)
        {
            UInt16 crc = 0xFFFF;
            byte[] result = new byte[2];

            for (int pos = 0; pos < (length - 2); pos++)
            {
                crc ^= (UInt16)message[pos];    //XOR

                for (int i = 8; i != 0; i--)
                {
                    // Passo ogni byte del pacchetto
                    if ((crc & 0x0001) != 0)
                    {
                        crc >>= 1;
                        crc ^= 0xA001;
                    }
                    else
                    {             
                        crc >>= 1;
                    }
                }
            }

            result[0] = (byte)(crc);        //LSB
            result[1] = (byte)(crc >> 8);   //MSB

            bool check = ((byte)(crc) == message[length - 2] && ((byte)(crc >> 8) == message[length - 1]));

            if (!check)
            {
                Console_print(" CRC Error - Expected: " + ((byte)(crc)).ToString("X").PadLeft(2,'0') + " " + ((byte)(crc >> 8)).ToString("X").PadLeft(2, '0') + " - Received: " + message[length - 2].ToString("X").PadLeft(2, '0') + " " + message[length - 1].ToString("X").PadLeft(2, '0'), null, 0);
            }

            return check;
        }

        public string timestamp()
        {
            return DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" +
                   DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" +
                   DateTime.Now.Second.ToString().PadLeft(2, '0');
        }

        // Cambia ordine inserimento righe nella RichTextBox di logi
        public void insertLogLinesAtTop(bool yes)
        {
            ordineTextBoxLog = yes;
        }

        //-------------------------------------------------------------------------------------
        //----------------Funzioni stampa su console o textBox array di byte-------------------
        //------------------------------------------------------------------------------------
        
        private void Console_printByte(String intestazione, byte[] query, int Length)
        {
            if (Length > 0)
            {
                String message = "";

                for (int i = 0; i < Length; i++)
                    message += query[i].ToString("X").PadLeft(2, '0') + " ";

                Console.WriteLine(intestazione + message);
            }
        }

        public string Console_print(string header, byte[] query, int Length)
        {
            if (Length > 0)
            {
                String message = "";
                String aa = "";

                for (int i = 0; i < Length; i++)
                {

                    aa = query[i].ToString("X");

                    if (aa.Length < 2)
                        aa = "0" + aa;

                    message += "" + aa + " ";
                }

                log.Enqueue(timestamp() + header + message + "\n");
                log2.Enqueue(timestamp() + header + message + "\n");

                return timestamp() + header + message + "\n";
            }
            else
            {
                if (header != null)
                {
                    if (header.Length > 0)
                    {
                        log.Enqueue(timestamp() + header + "\n");
                        log2.Enqueue(timestamp() + header + "\n");
                    }
                }

                return timestamp() + header + "\n";
            }
        }

        // Funzione equivalente alla precedente Application.DoEvents()
        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, new Action(delegate { }));
        }
    }

    public class ModBus_Item
    {
        public string Register { get; set; }
        public string Value { get; set; }
        public string ValueBin { get; set; }
        public string Notes { get; set; }
        public string Color { get; set; }
    }

    public class FixedSizedQueue<T>
    {
        ConcurrentQueue<T> q = new ConcurrentQueue<T>();
        private object lockObject = new object();

        public int Limit { get; set; }
        public void Enqueue(T obj)
        {
            q.Enqueue(obj);

            lock (lockObject)
            {
                T overflow;
                while (q.Count > Limit && q.TryDequeue(out overflow)) ;
            }
        }

        public bool TryDequeue(out T obj)
        {
            if(q.TryDequeue(out obj))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }

    public class ModbusException: Exception
    {
        public ModbusException() { }
        public ModbusException(string message) : base(message) { }
        public ModbusException(string message, Exception inner) : base(message, inner) { }
    };
}
