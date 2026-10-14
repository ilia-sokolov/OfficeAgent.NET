"""Start the bundled OfficeAgent MCP server with the arguments this command was given.

`officeagent-mcp --stdio` runs the server over standard input and output, as MCP clients
expect. On Linux and macOS the server replaces this process, so signals and exit codes reach
the client unchanged. Windows has no such replacement, so the server runs as a child that
shares this console, inside a job object that ends it when this launcher ends, however it
ends; its exit code is passed back.
"""
import os
import subprocess
import sys

from . import executable


def _end_with_this_process(process: "subprocess.Popen[bytes]") -> object:
    """Tie the server's lifetime to this launcher (Windows).

    A client that kills the launcher without closing its end of the stdio pipes would
    otherwise leave the server running. The job is closed when this process exits, and
    JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE then ends the server. Returns the job handle, which
    must stay open for the life of the launcher, or None when the job cannot be set up.
    """
    import ctypes
    from ctypes import wintypes

    class BasicLimits(ctypes.Structure):
        _fields_ = [
            ("PerProcessUserTimeLimit", ctypes.c_int64),
            ("PerJobUserTimeLimit", ctypes.c_int64),
            ("LimitFlags", wintypes.DWORD),
            ("MinimumWorkingSetSize", ctypes.c_size_t),
            ("MaximumWorkingSetSize", ctypes.c_size_t),
            ("ActiveProcessLimit", wintypes.DWORD),
            ("Affinity", ctypes.c_size_t),
            ("PriorityClass", wintypes.DWORD),
            ("SchedulingClass", wintypes.DWORD),
        ]

    class IoCounters(ctypes.Structure):
        _fields_ = [(name, ctypes.c_uint64) for name in (
            "ReadOperationCount", "WriteOperationCount", "OtherOperationCount",
            "ReadTransferCount", "WriteTransferCount", "OtherTransferCount")]

    class ExtendedLimits(ctypes.Structure):
        _fields_ = [
            ("BasicLimitInformation", BasicLimits),
            ("IoInfo", IoCounters),
            ("ProcessMemoryLimit", ctypes.c_size_t),
            ("JobMemoryLimit", ctypes.c_size_t),
            ("PeakProcessMemoryUsed", ctypes.c_size_t),
            ("PeakJobMemoryUsed", ctypes.c_size_t),
        ]

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateJobObjectW.restype = wintypes.HANDLE
    kernel32.CreateJobObjectW.argtypes = (wintypes.LPVOID, wintypes.LPCWSTR)
    kernel32.SetInformationJobObject.restype = wintypes.BOOL
    kernel32.SetInformationJobObject.argtypes = (wintypes.HANDLE, ctypes.c_int, wintypes.LPVOID, wintypes.DWORD)
    kernel32.AssignProcessToJobObject.restype = wintypes.BOOL
    kernel32.AssignProcessToJobObject.argtypes = (wintypes.HANDLE, wintypes.HANDLE)

    job = kernel32.CreateJobObjectW(None, None)
    if not job:
        return None
    limits = ExtendedLimits()
    limits.BasicLimitInformation.LimitFlags = 0x2000  # JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    extended_limit_information = 9  # JobObjectExtendedLimitInformation
    if not kernel32.SetInformationJobObject(job, extended_limit_information, ctypes.byref(limits), ctypes.sizeof(limits)):
        return None
    if not kernel32.AssignProcessToJobObject(job, int(process._handle)):  # noqa: SLF001 - Popen's own process handle
        return None
    return job


def main() -> int:
    path = executable()
    if not path.is_file():
        sys.stderr.write(
            f"officeagent-mcp: the bundled server is missing ({path}). Reinstall officeagent-mcp.\n"
        )
        return 1
    argv = [str(path), *sys.argv[1:]]

    if os.name == "nt":
        process = subprocess.Popen(argv)
        job = _end_with_this_process(process)  # noqa: F841 - held open until this launcher exits
        while True:
            try:
                return process.wait()
            except KeyboardInterrupt:
                # The server shares this console, so it receives Ctrl+C too; wait for it to stop.
                continue

    if not os.access(path, os.X_OK):
        # Installers keep the executable bit from the wheel; restore it if one did not.
        try:
            os.chmod(path, path.stat().st_mode | 0o111)
        except OSError:
            pass
    os.execv(str(path), argv)
    return 1  # execv does not return


if __name__ == "__main__":
    sys.exit(main())
