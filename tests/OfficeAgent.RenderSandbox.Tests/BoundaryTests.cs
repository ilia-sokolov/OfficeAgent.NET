using OfficeAgent.Abstractions;
using OfficeAgent.Deploy.Renderer;

namespace OfficeAgent.RenderSandbox.Tests;

/// <summary>
/// The container boundary a render runs inside. Each probe runs a shell command in a
/// container configured exactly as a render, from the same argument list.
/// </summary>
[Collection(SandboxCollection.Name)]
public sealed class BoundaryTests
{
    // /dev/tcp is bash's own socket support, so the probe needs no network tool in the image.
    private const string ConnectOut = "exec 3<>/dev/tcp/1.1.1.1/53 && echo connected";

    [Fact]
    public void Network_is_denied()
    {
        var sandboxed = Sandbox.Probe(ConnectOut);
        Assert.NotEqual(0, sandboxed.ExitCode);
        Assert.DoesNotContain("connected", sandboxed.Stdout);

        // Control: the same probe on Docker's default network connects, so a refusal above is
        // the sandbox and not a broken probe or an offline machine.
        var control = Sandbox.Probe(ConnectOut, withNetwork: true);
        Assert.True(control.ExitCode == 0 && control.Stdout.Contains("connected"),
            $"control probe could not connect, so the network test proves nothing: {control.Stderr}");
    }

    [Fact]
    public void Only_the_scratch_directory_is_writable()
    {
        var result = Sandbox.Probe(
            "for d in / /usr /etc /opt /opt/officeagent/worker /usr/lib/libreoffice/share/registry /var /root; do " +
            "  if touch \"$d/probe\" 2>/dev/null; then echo \"WRITABLE $d\"; fi; done; " +
            "touch /tmp/probe && echo scratch-ok");
        Assert.Equal(0, result.ExitCode);
        Assert.DoesNotContain("WRITABLE", result.Stdout);
        Assert.Contains("scratch-ok", result.Stdout);

        // Failed writes alone do not prove a read-only root: an unprivileged user is refused by
        // file permissions too. The mount itself must be read-only.
        var root = Sandbox.Probe("awk '$2 == \"/\" {print $4}' /proc/mounts").Stdout.Trim();
        Assert.StartsWith("ro", root.Split(',')[0]);
    }

    [Fact]
    public void The_worker_runs_unprivileged()
    {
        var result = Sandbox.Probe("id -u; id -g; grep -E '^(CapEff|CapBnd|NoNewPrivs):' /proc/self/status");
        var lines = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        Assert.Equal("10001", lines[0]);
        Assert.Equal("10001", lines[1]);
        Assert.Contains(lines, line => line.StartsWith("CapEff:") && line.EndsWith("0000000000000000"));
        // A non-root process has no effective capabilities anyway; dropping them all is what
        // empties the bounding set, the ceiling any escalation could reach.
        Assert.Contains(lines, line => line.StartsWith("CapBnd:") && line.EndsWith("0000000000000000"));
        Assert.Contains(lines, line => line.StartsWith("NoNewPrivs:") && line.EndsWith("1"));
    }

    [Fact]
    public void No_host_path_is_mounted()
    {
        // What Docker always provides; anything else would be a path the host shared.
        var allowed = new[] { "/", "/tmp", "/dev", "/dev/pts", "/dev/mqueue", "/dev/shm", "/dev/console",
                              "/etc/hosts", "/etc/hostname", "/etc/resolv.conf" };
        var result = Sandbox.Probe("awk '{print $2}' /proc/mounts");
        var unexpected = result.Stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(mount => !allowed.Contains(mount) && !mount.StartsWith("/proc") && !mount.StartsWith("/sys"))
            .ToList();
        Assert.True(unexpected.Count == 0, "unexpected mounts: " + string.Join(", ", unexpected));
    }

    // Forty background processes, then a marker. Bash aborts the script when a fork finally
    // fails, so the marker appears only if every fork succeeded.
    private const string ForkForty = "for ((i=0;i<40;i++)); do sleep 30 & done; echo all-forked";

    [Fact]
    public void The_process_limit_holds()
    {
        var limited = Sandbox.Probe(ForkForty, new SandboxLimits { Image = Sandbox.Reference, Pids = 32 });
        Assert.DoesNotContain("all-forked", limited.Stdout);
        Assert.Contains("fork: Resource temporarily unavailable", limited.Stderr);

        // Control: the same script under a limit it fits in completes, so the refusal above is
        // the limit and not the script.
        var control = Sandbox.Probe(ForkForty, new SandboxLimits { Image = Sandbox.Reference, Pids = 256 });
        Assert.Contains("all-forked", control.Stdout);
    }

    // Holds 300 MiB in one shell variable, then prints a marker.
    private const string Allocate = "v=$(head -c 300m /dev/zero | tr '\\0' a); echo survived ${#v}";

    [Fact]
    public void The_memory_limit_is_enforced_by_the_kernel()
    {
        var limited = Sandbox.Probe(Allocate, new SandboxLimits { Image = Sandbox.Reference, MemoryBytes = 128L * 1024 * 1024 });
        Assert.DoesNotContain("survived", limited.Stdout);
        Assert.Equal(137, limited.ExitCode); // killed by the kernel's OOM killer inside the cgroup

        var control = Sandbox.Probe(Allocate, new SandboxLimits { Image = Sandbox.Reference, MemoryBytes = 1024L * 1024 * 1024 });
        Assert.Contains("survived", control.Stdout);
    }

    [Fact]
    public void Scratch_space_is_bounded()
    {
        var result = Sandbox.Probe("dd if=/dev/zero of=/tmp/fill bs=1M count=100 2>&1; echo done",
            new SandboxLimits { Image = Sandbox.Reference, ScratchBytes = 16L * 1024 * 1024 });
        Assert.Contains("No space left on device", result.Stdout);
    }

    [Fact]
    public async Task A_timed_out_render_leaves_no_container_behind()
    {
        Assert.Equal(0, Sandbox.ContainersNamed("officeagent-render-"));
        var renderer = Sandbox.Renderer(new SandboxLimits { Image = Sandbox.Reference, StartupAllowance = TimeSpan.Zero });

        var result = await renderer.RenderAsync(new MemoryStream(Documents.Word(30)), new RenderOptions
        {
            FileName = "long.docx",
            Timeout = TimeSpan.FromMilliseconds(500)
        });

        Assert.False(result.Succeeded);
        Assert.Equal(RenderFailureCodes.RenderTimeout, result.FailureCode);
        // The container, and with it every process LibreOffice started, is gone.
        Assert.Equal(0, Sandbox.ContainersNamed("officeagent-render-"));
    }
}
