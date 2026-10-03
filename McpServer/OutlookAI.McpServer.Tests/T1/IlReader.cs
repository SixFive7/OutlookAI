using System.Reflection;
using System.Reflection.Emit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>One instruction of a method body that the static checks care about.</summary>
/// <param name="Op">The opcode.</param>
/// <param name="Method">The method a call, callvirt, newobj, ldftn or ldvirtftn names; null otherwise or when it cannot be resolved.</param>
/// <param name="Text">The string an ldstr pushes; null otherwise.</param>
internal readonly record struct IlInstruction(OpCode Op, MethodBase? Method, string? Text);

/// <summary>
/// A real opcode walk over compiled method bodies, for the T1 pins that have to read what the code
/// DOES rather than what it is called: which methods it calls, which delegates it creates, which
/// strings it pushes. A walk rather than a byte search, because a byte search reads operand bytes as
/// opcodes, and a pin built on that can be green by accident.
/// </summary>
internal static class IlReader
{
    private const BindingFlags Everything = BindingFlags.Public | BindingFlags.NonPublic
        | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;

    /// <summary>Every opcode, by its value, so the walk can step over operands of every size.</summary>
    private static readonly Dictionary<ushort, OpCode> OpCodesByValue = typeof(OpCodes)
        .GetFields(BindingFlags.Public | BindingFlags.Static)
        .Select(f => (OpCode)f.GetValue(null)!)
        .ToDictionary(op => unchecked((ushort)op.Value));

    /// <summary>
    /// The instructions of <paramref name="method"/> that name a method or push a string, in order.
    /// A bodiless method (abstract, extern, an interface member) has none.
    /// </summary>
    public static IReadOnlyList<IlInstruction> Read(MethodBase method)
    {
        byte[]? il;
        try
        {
            il = method.GetMethodBody()?.GetILAsByteArray();
        }
        catch (Exception ex) when (ex is InvalidOperationException or NotSupportedException or BadImageFormatException)
        {
            il = null;
        }

        if (il == null)
        {
            return Array.Empty<IlInstruction>();
        }

        Type[]? typeArguments = method.DeclaringType is { IsGenericType: true } declaring ? declaring.GetGenericArguments() : null;
        Type[]? methodArguments = method is MethodInfo { IsGenericMethod: true } generic ? generic.GetGenericArguments() : null;

        List<IlInstruction> found = new();
        int position = 0;
        while (position < il.Length)
        {
            ushort value = il[position++];
            if (value == 0xFE)
            {
                value = (ushort)(0xFE00 | il[position++]);
            }

            OpCode op = OpCodesByValue[value];
            int operand = position;
            position += OperandSize(op.OperandType, il, operand);

            if (op == OpCodes.Ldstr)
            {
                string? text;
                try
                {
                    text = method.Module.ResolveString(BitConverter.ToInt32(il, operand));
                }
                catch (ArgumentException)
                {
                    text = null;
                }

                found.Add(new IlInstruction(op, null, text));
            }
            else if (op.OperandType == OperandType.InlineMethod)
            {
                MethodBase? target;
                try
                {
                    target = method.Module.ResolveMethod(BitConverter.ToInt32(il, operand), typeArguments, methodArguments);
                }
                catch (Exception ex) when (ex is ArgumentException or BadImageFormatException or TypeLoadException)
                {
                    target = null;
                }

                found.Add(new IlInstruction(op, target, null));
            }
        }

        return found;
    }

    /// <summary>Every method, constructor and type initializer <paramref name="type"/> declares itself.</summary>
    public static IEnumerable<MethodBase> BodiesOf(Type type)
    {
        IEnumerable<MethodBase> methods = type.GetMethods(Everything).Cast<MethodBase>().Concat(type.GetConstructors(Everything));
        return type.TypeInitializer == null ? methods : methods.Append(type.TypeInitializer);
    }

    /// <summary>
    /// Every string <paramref name="type"/>'s compiled code can push, its compiler-generated nested
    /// types included - async and iterator state machines and lambda display classes, which is where
    /// the body of most test methods actually lives.
    /// </summary>
    public static HashSet<string> StringsOf(Type type)
    {
        HashSet<string> strings = new(StringComparer.Ordinal);
        foreach (Type each in WithNestedTypes(type))
        {
            foreach (MethodBase body in BodiesOf(each))
            {
                foreach (IlInstruction instruction in Read(body))
                {
                    if (instruction.Text != null)
                    {
                        strings.Add(instruction.Text);
                    }
                }
            }
        }

        return strings;
    }

    /// <summary>Every method <paramref name="type"/>'s compiled code names, nested types included.</summary>
    public static IEnumerable<MethodBase> MethodsNamedBy(Type type)
    {
        return WithNestedTypes(type)
            .SelectMany(BodiesOf)
            .SelectMany(Read)
            .Where(i => i.Method != null)
            .Select(i => i.Method!);
    }

    /// <summary><paramref name="type"/> and every type nested in it, at any depth.</summary>
    public static IEnumerable<Type> WithNestedTypes(Type type)
    {
        yield return type;
        foreach (Type nested in type.GetNestedTypes(BindingFlags.Public | BindingFlags.NonPublic))
        {
            foreach (Type inner in WithNestedTypes(nested))
            {
                yield return inner;
            }
        }
    }

    private static int OperandSize(OperandType type, byte[] il, int operand)
    {
        switch (type)
        {
            case OperandType.InlineNone:
                return 0;
            case OperandType.ShortInlineBrTarget:
            case OperandType.ShortInlineI:
            case OperandType.ShortInlineVar:
                return 1;
            case OperandType.InlineVar:
                return 2;
            case OperandType.InlineI8:
            case OperandType.InlineR:
                return 8;
            case OperandType.InlineSwitch:
                return 4 + (4 * BitConverter.ToInt32(il, operand));
            default:
                return 4;
        }
    }
}
