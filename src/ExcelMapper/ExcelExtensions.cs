using System;
using System.Diagnostics;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelMapper
{

    public static class ExcelExtensions
    {
        public static HeaderConfiguration<T> Column<T>(this HeaderConfiguration<T> configuration,
            string header,
            Expression<Func<T, object>> prop)  where T:new()
        {
            configuration.Add(prop, header);
            return configuration;
        }

        public static HeaderConfiguration<T> Column<T>(this HeaderConfiguration<T> configuration,
            Expression<Func<T, object>> prop) where T:new() =>
            configuration.Column(null, prop);

        public static ColumnConfiguration<T> Column<T>(this ColumnConfiguration<T> configuration,
            int column,
            Expression<Func<T, object>> prop
        )  where T:new()
        {
            configuration.Add(prop, column);
            return configuration;
        }

        public static Settings Worksheet<T>(this Settings settings, string sheet,
            Action<HeaderConfiguration<T>> configure)  where T:new()
        {
            var configuration = new HeaderConfiguration<T>();
            configure(configuration);
            settings.Add(configuration, sheet);
            return settings;
        }

        public static Settings Worksheet<T>(this Settings settings, Action<HeaderConfiguration<T>> configure)  where T:new()
            => settings.Worksheet(null, configure);

        public static Settings Worksheet<T>(this Settings settings, string sheet,
            Action<ColumnConfiguration<T>> configure) where T:new()
        {
            var configuration = new ColumnConfiguration<T>();
            configure(configuration);
            settings.Add(configuration, sheet);
            return settings;
        }

        public static Settings Worksheet<T>(this Settings settings, Action<ColumnConfiguration<T>> configure)  where T:new()
            => settings.Worksheet(null, configure);

        internal static MemberInfo GetMemberFromExpression<T>(this Expression<Func<T, object>> field)
        {
            var member = field.Body as MemberExpression ??
                         (field.Body as UnaryExpression)?.Operand as MemberExpression;
            if (member == null)
                throw new ArgumentException("invalid member selector");

            Debug.WriteLine($"{typeof(T).Name}.{member.Member.Name}:{member.Type.Name}");

            return member.Member;
        }

        internal static void Set(this object instance , MemberInfo member,object value)
        {
            object ProcessValue(Type expectedType)
            {
                if (expectedType == typeof(string))
                    return value?.ToString();
                
                if (expectedType != value.GetType())
                    return Convert.ChangeType(value, expectedType);
                
                return value;
            }
            
            switch (member)
            {
                case PropertyInfo p:
                    p.SetValue(instance, ProcessValue(p.PropertyType));
                    break;
                case FieldInfo f:
                    f.SetValue(instance, ProcessValue(f.FieldType));
                    break;
                default: throw new InvalidCastException($"can't set value of {member.DeclaringType}.{member.Name} to {value}");
            }
        }

    }
}