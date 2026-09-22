using System.Diagnostics;
using System.Text;
using System.Text.Json;
using OfficeAgent.Deploy.Renderer;

namespace OfficeAgent.RenderSandbox.Tests;

/// <summary>Builds the reference and control images once, and runs things inside them.</summary>
public sealed class Sandbox
{
    public const string Reference = "officeagent-renderer:reference";
    public const string PermissiveControl = "officeagent-renderer:permissive-control";

    public Sandbox()
    {
        var engine = Run("docker", "version", "--format", "{{.Server.Os}}");
        if (engine.ExitCode != 0 || engine.Stdout.Trim() != "linux")
            throw new InvalidOperationException(
                "These reference tests need a Linux container engine (docker version reported: " +
                $"'{engine.Stdout.Trim()}', exit {engine.ExitCode}). They fail rather than skip.");

        if (Environment.GetEnvironmentVariable("OFFICEAGENT_SANDBOX_SKIP_BUILD") == "1") return;
        var root = RepositoryRoot();
        Require(Run("dotnet", "publish", Path.Combine(root, "deploy", "renderer", "worker"), "-c", "Release",
            "-o", Path.Combine(root, "deploy", "renderer", "worker", "bin", "publish"), "-nologo"), "publish the worker");
        var context = Path.Combine(root, "deploy", "renderer");
        Require(Run("docker", "build", "--target", "hardened", "-t", Reference, context), "build the reference image");
        Require(Run("docker", "build", "--target", "permissive", "-t", PermissiveControl, context), "build the control image");
    }

    /// <summary>The reference renderer with the deployment recipe's limits.</summary>
    public static SandboxedDocumentRenderer Renderer(SandboxLimits? limits = null) =>
        new(limits ?? new SandboxLimits { Image = Reference });

    /// <summary>
    /// Runs a shell command in a container configured exactly as a render, so the boundary a
    /// render gets is the boundary this probes.
    /// </summary>
    public static Result Probe(string script, SandboxLimits? limits = null, bool withNetwork = false)
    {
        var arguments = Renderer(limits).RunArguments("officeagent-probe-" + Guid.NewGuid().ToString("N")).ToList();
        if (withNetwork)
        {
            // The control for the network test: the same container with Docker's default network.
            var index = arguments.IndexOf("--network");
            arguments.RemoveRange(index, 2);
        }
        var image = arguments[^1];
        arguments.RemoveAt(arguments.Count - 1);
        arguments.AddRange(new[] { "--entrypoint", "bash", image, "-c", script });
        return Run("docker", arguments.ToArray());
    }

    /// <summary>Speaks the worker protocol directly, to read the audit a host never sees.</summary>
    public static JsonElement Worker(string image, byte[] document, string fileName, bool audit)
    {
        var limits = new SandboxLimits
        {
            Image = image,
            Environment = audit ? new Dictionary<string, string> { ["OFFICEAGENT_SANDBOX_AUDIT"] = "1" } : new()
        };
        var arguments = Renderer(limits).RunArguments("officeagent-audit-" + Guid.NewGuid().ToString("N"));
        var header = JsonSerializer.Serialize(new
        {
            fileName, dpi = 72, timeoutMilliseconds = 120_000, maximumInputBytes = 50_000_000L,
            maximumWorkingSetBytes = 900_000_000L, maximumPages = 50, maximumOutputBytes = 100_000_000L
        });
        var input = Encoding.UTF8.GetBytes(header + "\n").Concat(document).ToArray();
        var result = Run("docker", arguments.ToArray(), input);
        Require(result, "run the worker");
        return JsonDocument.Parse(result.Stdout).RootElement.Clone();
    }

    public static int ContainersNamed(string prefix)
    {
        var listed = Run("docker", "ps", "-a", "--filter", "name=" + prefix, "--format", "{{.Names}}");
        return listed.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries).Length;
    }

    public sealed record Result(int ExitCode, string Stdout, string Stderr);

    public static Result Run(string executable, params string[] arguments) => Run(executable, arguments, null);

    public static Result Run(string executable, string[] arguments, byte[]? stdin)
    {
        var start = new ProcessStartInfo(executable)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            RedirectStandardInput = stdin is not null
        };
        foreach (var argument in arguments) start.ArgumentList.Add(argument);
        using var process = Process.Start(start)!;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        if (stdin is not null)
        {
            process.StandardInput.BaseStream.Write(stdin);
            process.StandardInput.Close();
        }
        if (!process.WaitForExit(600_000))
        {
            process.Kill(entireProcessTree: true);
            throw new TimeoutException($"{executable} {string.Join(' ', arguments)} did not finish.");
        }
        return new Result(process.ExitCode, stdout.Result, stderr.Result);
    }

    private static void Require(Result result, string what)
    {
        if (result.ExitCode != 0)
            throw new InvalidOperationException($"Could not {what} (exit {result.ExitCode}): {result.Stderr}{result.Stdout}");
    }

    public static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "OfficeAgent.NET.sln")))
            directory = directory.Parent;
        return directory?.FullName ?? throw new DirectoryNotFoundException("repository root not found");
    }
}

[CollectionDefinition(Name)]
public sealed class SandboxCollection : ICollectionFixture<Sandbox>
{
    public const string Name = "sandbox";
}
