using BlazorDownloadFile;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelWizard.Service;

public class FakeBlazorDownloadFileService : IBlazorDownloadFileService
{
    public ValueTask AddBuffer(string bytesBase64)
    {
        throw new InvalidOperationException("You can only invoke this method in Blazor context");
    }

    public ValueTask AddBuffer(string bytesBase64, CancellationToken cancellationToken)
    {
        throw new InvalidOperationException("You can invoke this method in Blazor context");
    }

    public ValueTask AddBuffer(string bytesBase64, TimeSpan timeOut)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask AddBuffer(byte[] bytes)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask AddBuffer(byte[] bytes, CancellationToken cancellationToken)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask AddBuffer(byte[] bytes, TimeSpan timeOut)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask AddBuffer(IEnumerable<byte> bytes)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask AddBuffer(IEnumerable<byte> bytes, CancellationToken cancellationToken)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask AddBuffer(IEnumerable<byte> bytes, TimeSpan timeOut)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask AddBuffer(Stream stream)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask AddBuffer(Stream stream, CancellationToken cancellationToken)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask AddBuffer(Stream stream, CancellationToken streamReadcancellationToken, TimeSpan timeOutJavaScript)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<bool> AnyBuffer()
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<int> BuffersCount()
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask ClearBuffers()
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadBase64Buffers(string fileName, string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadBase64Buffers(string fileName, CancellationToken cancellationToken,
        string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadBase64Buffers(string fileName, TimeSpan timeOut, string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadBinaryBuffers(string fileName, string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadBinaryBuffers(string fileName, CancellationToken cancellationToken,
        string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadBinaryBuffers(string fileName, TimeSpan timeOut, string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, string bytesBase64, string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, string bytesBase64, CancellationToken cancellationToken,
        string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, string bytesBase64, TimeSpan timeOut,
        string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, byte[] bytes, string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, byte[] bytes, CancellationToken cancellationToken,
        string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, byte[] bytes, TimeSpan timeOut,
        string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, IEnumerable<byte> bytes, string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, IEnumerable<byte> bytes, CancellationToken cancellationToken,
        string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, IEnumerable<byte> bytes, TimeSpan timeOut,
        string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, Stream stream, string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, Stream stream, CancellationToken cancellationTokenBytesRead,
        CancellationToken cancellationTokenJavaScriptInterop, string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, Stream stream, CancellationToken cancellationTokenBytesRead,
        TimeSpan timeOutJavaScriptInterop, string contentType = "application/octet-stream")
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFileFromText(string fileName, string plainText, Encoding encoding, string contentType = "text/plain",
        bool encoderShouldEmitIdentifier = false)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFileFromText(string fileName, string plainText, Encoding encoding,
        CancellationToken cancellationToken, string contentType = "text/plain", bool encoderShouldEmitIdentifier = false)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFileFromText(string fileName, string plainText, Encoding encoding, TimeSpan timeOut,
        string contentType = "text/plain", bool encoderShouldEmitIdentifier = false)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, string bytesBase64, int bufferSize = 32768,
        string contentType = "application/octet-stream", IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, string bytesBase64, CancellationToken cancellationToken, int bufferSize = 32768,
        string contentType = "application/octet-stream", IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, string bytesBase64, TimeSpan timeOut, int bufferSize = 32768,
        string contentType = "application/octet-stream", IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, byte[] bytes, int bufferSize = 32768,
        string contentType = "application/octet-stream", IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, byte[] bytes, CancellationToken cancellationToken, int bufferSize = 32768,
        string contentType = "application/octet-stream", IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, byte[] bytes, TimeSpan timeOut, int bufferSize = 32768,
        string contentType = "application/octet-stream", IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, IEnumerable<byte> bytes, int bufferSize = 32768,
        string contentType = "application/octet-stream", IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, IEnumerable<byte> bytes, CancellationToken cancellationToken, int bufferSize = 32768,
        string contentType = "application/octet-stream", IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, IEnumerable<byte> bytes, TimeSpan timeOut, int bufferSize = 32768,
        string contentType = "application/octet-stream", IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, Stream stream, int bufferSize = 32768,
        string contentType = "application/octet-stream", IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, Stream stream, CancellationToken cancellationTokenBytesRead,
        CancellationToken cancellationTokenJavaScriptInterop, int bufferSize = 32768,
        string contentType = "application/octet-stream", IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFile(string fileName, Stream stream, CancellationToken cancellationTokenBytesRead,
        TimeSpan timeOutJavaScriptInterop, int bufferSize = 32768, string contentType = "application/octet-stream",
        IProgress<double>? progress = null)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFileFromText(string fileName, string plainText, Encoding encoding, int bufferSize = 32768,
        string contentType = "text/plain", IProgress<double>? progress = null, bool encoderShouldEmitIdentifier = false)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFileFromText(string fileName, string plainText, Encoding encoding,
        CancellationToken cancellationToken, int bufferSize = 32768, string contentType = "text/plain",
        IProgress<double>? progress = null, bool encoderShouldEmitIdentifier = false)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }

    public ValueTask<DownloadFileResult> DownloadFileFromText(string fileName, string plainText, Encoding encoding, TimeSpan timeOut,
        int bufferSize = 32768, string contentType = "text/plain", IProgress<double>? progress = null,
        bool encoderShouldEmitIdentifier = false)
    {
        throw new InvalidOperationException("You cannot invoke method in not Blazor context");
    }
}