// THE FEW TYPES NEWER C# NEEDS AND .NET FRAMEWORK 4.8 DOES NOT HAVE (Q172, 2026-11-18).
//
// The add-in compiles as C# 15 since the maintainer moved OutlookAI to .NET 11; it stays on
// .NET Framework 4.8 because a VSTO add-in cannot run on modern .NET. Most of what C# added since
// 7.3 is syntax and needs nothing from the runtime. A few features find a type by its name in the
// compilation instead, and .NET Framework 4.8 ships none of these, so they are declared here,
// internal, as source - no package (AGENTS.md, Dependencies). The compiler uses them where it
// finds them; nothing calls them.
//
//   * IsExternalInit: init accessors and records (C# 9).
//   * RequiredMemberAttribute, CompilerFeatureRequiredAttribute, SetsRequiredMembersAttribute:
//     required members (C# 11).
//   * CallerArgumentExpressionAttribute: the argument's own text in a guard's message (C# 10).
//
// What needs the RUNTIME itself, the compiler refuses on .NET Framework 4.8 whatever is declared
// here - default interface members (CS8701), static abstract interface members (CS8919),
// covariant returns (CS8830) and the like - and spans need System.Memory, which the add-in does
// not reference. So: no feature that needs the modern runtime.
//
// Only the add-in compiles this file. A file the add-in shares with OutlookAI.Core or the test
// project compiles there too, where Core's net48 target has none of these types, so a shared file
// uses none of these features unless the projects that compile it link this file as well (Q154:
// one copy).

namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit
    {
    }

    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct | AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    internal sealed class RequiredMemberAttribute : Attribute
    {
    }

    [AttributeUsage(AttributeTargets.All, AllowMultiple = true, Inherited = false)]
    internal sealed class CompilerFeatureRequiredAttribute : Attribute
    {
        public const string RefStructs = "RefStructs";
        public const string RequiredMembers = "RequiredMembers";

        public CompilerFeatureRequiredAttribute(string featureName)
        {
            FeatureName = featureName;
        }

        public string FeatureName { get; }

        public bool IsOptional { get; set; }
    }

    [AttributeUsage(AttributeTargets.Parameter, AllowMultiple = false, Inherited = false)]
    internal sealed class CallerArgumentExpressionAttribute : Attribute
    {
        public CallerArgumentExpressionAttribute(string parameterName)
        {
            ParameterName = parameterName;
        }

        public string ParameterName { get; }
    }
}

namespace System.Diagnostics.CodeAnalysis
{
    [AttributeUsage(AttributeTargets.Constructor, AllowMultiple = false, Inherited = false)]
    internal sealed class SetsRequiredMembersAttribute : Attribute
    {
    }
}
