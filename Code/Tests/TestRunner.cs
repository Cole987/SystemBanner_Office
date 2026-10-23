//Office add-in: Copyright (c) 2026, Cole. MIT License Applies.
using System;
using System.Collections.Generic;
using System.Reflection;

namespace SystemBannerOffice.Tests
{
    //a small test runner with no dependency, so the tests build with the .NET 2.0 / 3.5 compilers of Windows
    //like the add-in: every public static void method of a class marked [TestClass] is a test
    [AttributeUsage(AttributeTargets.Class)]
    public class TestClassAttribute : Attribute { }

    public static class Assert
    {
        public static void AreEqual(object expected, object actual)
        {
            if (!object.Equals(expected, actual))
            {
                throw new Exception("expected <" + expected + ">\n       but was  <" + actual + ">");
            }
        }

        public static void IsTrue(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception(message);
            }
        }

        public static void IsNull(object value)
        {
            if (value != null)
            {
                throw new Exception("expected null but was <" + value + ">");
            }
        }

        public static void IsNotNull(object value, string message)
        {
            if (value == null)
            {
                throw new Exception(message);
            }
        }
    }

    public static class TestRunner
    {
        public static int Main(string[] args)
        {
            int passed = 0;
            List<string> failures = new List<string>();
            foreach (Type type in Assembly.GetExecutingAssembly().GetTypes())
            {
                if (type.GetCustomAttributes(typeof(TestClassAttribute), false).Length == 0)
                {
                    continue;
                }
                foreach (MethodInfo method in type.GetMethods(BindingFlags.Public | BindingFlags.Static))
                {
                    if (method.GetParameters().Length != 0 || method.ReturnType != typeof(void))
                    {
                        continue;
                    }
                    string name = type.Name + "." + method.Name;
                    if (args.Length > 0 && name.IndexOf(args[0], StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }
                    try
                    {
                        method.Invoke(null, null);
                        passed++;
                    }
                    catch (TargetInvocationException ex)
                    {
                        failures.Add(name + "\n       " + ex.InnerException.Message);
                    }
                }
            }
            foreach (string failure in failures)
            {
                Console.WriteLine("FAIL " + failure);
            }
            Console.WriteLine(passed + " passed, " + failures.Count + " failed");
            return failures.Count == 0 ? 0 : 1;
        }
    }
}
