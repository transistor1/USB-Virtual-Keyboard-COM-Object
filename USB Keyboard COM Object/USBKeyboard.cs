using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.IO.Ports;
using System.Text.RegularExpressions;
using System.Reflection;
using System.IO;
using System.Management;

namespace FieldEffect
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("68038726-ABB4-4C1E-B455-519F0648EE93")]
    public interface IUSBKeyboard
    {
        /// <summary>
        /// Connect to the named serial port, at 256000 baud.
        /// </summary>
        /// <param name="COMPort">The named serial port: COM1, COM2, etc.</param>
        void Connect(string COMPort);

        /// <summary>
        /// Load a text map from memory.
        /// </summary>
        /// <param name="TextMapAsString">The text map to load</param>
        void LoadTextMap(string TextMapAsString);

        /// <summary>
        /// Send a key to the connected port.
        /// </summary>
        /// <param name="KeyChar">A key char specified on the left hand side of a TextMap expression</param>
        void SendKey(string KeyChar);

        /// <summary>
        /// Send a 3-byte key code
        /// </summary>
        /// <param name="Byte1">First byte</param>
        /// <param name="Byte2">Second byte</param>
        /// <param name="Byte3">Third byte</param>
        void SendKeyCode(int Byte1, int Byte2, int Byte3);

        /// <summary>
        /// Send a string of characters, excluding special characters or hex
        /// </summary>
        /// <param name="Text"></param>
        void SendString(string Text);

        /// <summary>
        /// Disconnect from the COM port
        /// </summary>
        void Disconnect();

        /// <summary>
        /// Get a delimited list of COM port names
        /// </summary>
        /// <returns></returns>
        string GetCOMPortNames();

    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IUSBKeyboard))]
    [Guid("CD289F80-F36C-47B7-AE06-42DD0FB59D43")]
    public sealed class USBKeyboard : IUSBKeyboard
    {
        //The USB Virtual Keyboard only connects at 256000
        const int COM_BAUD_RATE = 256000;
        private SerialPort Port;
        private Dictionary<string, byte[]> KeyMap = new Dictionary<string, byte[]>();
        //private Dictionary<string, byte[]> SpecialChars = new Dictionary<string, byte[]>();
        
        public void Connect(string COMPort)
        {
            this.LoadDefaultTextMap();
            this.Port = new SerialPort(COMPort, COM_BAUD_RATE, Parity.None, 8, StopBits.One);
            this.Port.Open();
            
        }

        public void Disconnect()
        {
            this.Port.Close();
            this.Port.Dispose();
        }

        public void LoadTextMap(string TextMapAsString)
        {
            //Regex rx = new Regex(@"(.+).+\=\>.+\[(.+)\]");
            Regex rx = new Regex(@"(?:[""']){0,1} ([^""'\s]+) (?:[""']){0,1} .+ \=\> .+ \[(.+)\]", RegexOptions.IgnorePatternWhitespace);

            //Clear the existing key map
            this.KeyMap.Clear();

            foreach (string line in TextMapAsString.Split(new string[]{"\r\n"}, StringSplitOptions.None))
            {
                MatchCollection mc = rx.Matches(line);
                //Was there a match?
                if (mc.Count == 1)
                {
                    if (mc[0].Groups.Count == 3)
                    {
                        string CharacterName = mc[0].Groups[1].Value;
                        string CharacterBytes = mc[0].Groups[2].Value;

                        List<byte> Bytes = new List<byte>();
                        
                        foreach (string ByteString in CharacterBytes.Split(new char[] {','}))
                        {
                            byte result = 0;

                            try
                            {
                                result = byte.Parse(ByteString.Substring(2), System.Globalization.NumberStyles.HexNumber);
                                Bytes.Add(result);
                            }
                            catch (Exception e)
                            {
                                this.KeyMap.Clear();
                                throw new FormatException("Couldn't parse TextMap.", e);
                            }
                        }

                        if (CharacterName.Length > 2)
                        {
                            if (CharacterName.Substring(0, 2).ToLower() == @"\x")
                            {
                                try
                                {
                                    int val = int.Parse(CharacterName.Substring(2), System.Globalization.NumberStyles.HexNumber);
                                    CharacterName = "" + (char)val;
                                }
                                catch (Exception e)
                                {
                                    this.KeyMap.Clear();
                                    throw new FormatException("Couldn't parse TextMap.", e);
                                }
                            }
                            else if (CharacterName.Substring(0, 2).ToLower() == @"\s")
                            {
                                CharacterName = CharacterName.ToLower();
                            }
                        }

                        this.KeyMap.Add(CharacterName, Bytes.ToArray());
                    }
                }
            }
        }

        public void SendKey(string KeyChar)
        {

            if (KeyChar.Length > 2)
            {
                if (KeyChar.Substring(0, 2).ToLower() == @"\x")
                {
                    try
                    {
                        int val = int.Parse(KeyChar.Substring(2), System.Globalization.NumberStyles.HexNumber);
                        KeyChar = "" + (char)val;
                    }
                    catch (Exception e)
                    {
                        this.KeyMap.Clear();
                        throw new FormatException(string.Format("Illegal character {0}", KeyChar), e);
                    }
                }
                else if (KeyChar.Substring(0, 2).ToLower() == @"\s")
                {
                    KeyChar = KeyChar.ToLower();
                }
            }

            if (this.KeyMap.Keys.Contains(KeyChar))
            {
                SendBytes(this.KeyMap[KeyChar]);
            }
            else if (this.KeyMap.Keys.Contains(string.Format("{0}",KeyChar)))
            {
                SendBytes(this.KeyMap[string.Format("{0}", KeyChar)]);
            }
            else
            {
                throw new ArgumentException(string.Format("The character {0} (hex {1:X}) does not exist in your keymap.", KeyChar, (int)KeyChar[0]));
            }
        }

        public void SendString(string Text)
        {
            for (int i = 0; i < Text.Length; i++)
            {
                SendKey(Text[i].ToString());
            }
        }

        public void SendKeyCode(int Byte1, int Byte2, int Byte3)
        {
            this.SendBytes(new byte[] { (byte)Byte1, (byte)Byte2, (byte)Byte3 });
        }

        public string GetCOMPortNames()
        {
            return string.Join("|", SerialPort.GetPortNames());
        }

        private void LoadDefaultTextMap()
        {
            using (Stream DefaultTextMapStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("FieldEffect.TextMap.txt"))
            using (StreamReader TextMapReader = new StreamReader(DefaultTextMapStream))
            {
                this.LoadTextMap(TextMapReader.ReadToEnd());
            }
        }

        private void SendBytes(byte[] bytes)
        {
            this.Port.Write(bytes, 0, bytes.Length);

            //Read back all the results until we get a null.
            //This [hopefully] ensures that our character has been
            //sent before we continue.
            while (this.Port.ReadChar() != 0) ;
        }
    }
}
