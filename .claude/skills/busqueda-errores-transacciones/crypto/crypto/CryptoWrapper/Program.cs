using System;
using SecurityUtils;
using SecurityUtils.Entity;

// CLI puente para exponer SecurityUtils.AesCryptoPayload (AES-256-GCM con key
// derivada por SHA-256, tag de 128 bits) a procesos no-.NET (ej. el script
// Python de esta skill), sin necesidad de pythonnet.
//
// Uso:
//   dotnet CryptoWrapper.dll encrypt   < (linea1=key, linea2=texto plano)  -> stdout: payload Base64
//   dotnet CryptoWrapper.dll decrypt   < (linea1=key, linea2=payload)      -> stdout: texto plano
//
// La key y el texto SIEMPRE se pasan por stdin (nunca por argumentos de
// linea de comandos), para que no queden visibles en el listado de procesos
// del sistema operativo.

if (args.Length != 1 || (args[0] != "encrypt" && args[0] != "decrypt"))
{
    Console.Error.WriteLine("Uso: CryptoWrapper <encrypt|decrypt>  (key y texto por stdin, una linea cada uno)");
    Environment.Exit(2);
}

string? key = Console.ReadLine();
string? text = Console.ReadLine();

if (string.IsNullOrEmpty(key) || text is null)
{
    Console.Error.WriteLine("Se esperaban 2 lineas por stdin: key y texto.");
    Environment.Exit(2);
}

try
{
    if (args[0] == "encrypt")
    {
        string payload = AesCryptoPayload.EncryptGcmNoPadding_256_WithSha256DerivedKey(
            text, key, Constants.CryptoEncoding.Base64);
        Console.Out.Write(payload);
    }
    else
    {
        string plain = AesCryptoPayload.DecryptGcmNoPadding_256_WithSha256DerivedKey(
            text, key, Constants.CryptoEncoding.Base64);
        Console.Out.Write(plain);
    }
}
catch (Exception ex)
{
    // Nunca imprimir key/texto en el error; solo el tipo de excepcion.
    Console.Error.WriteLine($"ERROR:{ex.GetType().Name}");
    Environment.Exit(1);
}
