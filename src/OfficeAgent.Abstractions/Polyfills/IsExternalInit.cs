#if NETSTANDARD2_0
namespace System.Runtime.CompilerServices;

/// <summary>
/// Polyfill for the C# 9 <c>init</c> accessor on <c>netstandard2.0</c>. The compiler
/// requires this type to exist when an assembly uses <c>init</c>; on net5+ it ships in
/// the BCL, on <c>netstandard2.0</c> we supply it ourselves.
/// </summary>
internal static class IsExternalInit { }
#endif
