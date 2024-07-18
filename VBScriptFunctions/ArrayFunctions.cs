using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VBScriptFunctions
{
    public class ArrayFunctions
    {

        /// <summary>
        ///  Функция Array возвращает массив
        /// </summary>
        /// <remarks>
        ///  The Array function returns a variant containing an array.
        /// </remarks>
        /// 
        /// 
        /// <param name="values">Разделенные запятой элементы массива</param>
        /// <returns>Array</returns>
        public static Object[] Array(params object[] values)
        {
            return values.ToArray();
        }

        /// <summary>
        ///  Функция Filter возвращает массив из значений отфильтрованных на основе критерии фильтрации.
        /// </summary>
        /// <remarks>
        ///  The Filter function returns a zero-based array that contains a subset of a string array based on a filter criteria.
        /// </remarks>
        /// 
        /// <param name="compare">0 - учет регистра, 1 - без учета регистра</param>
        /// <param name="include">Bool: true - вернет массив, который содержит искомую строку; flase - вернет массив из значений, которые не содержат строку</param>
        /// <returns>Array</returns>
        /// <exception cref="ArgumentException"></exception>
        public static Object[] Filter(object[] inputstrings, object value, object include, object compare)
        {
            if (!((value is int || value is long || value is bool || value is double || value is float || value is string)
                && (include is int || include is long || include is bool || include is double || include is float || include is string)
            && (compare is int || compare is string || compare is float || compare is double)))
                throw new ArgumentException();
            List<object> result = new List<object>();
            double includeToDoulble = Convert.ToDouble(include);//Если в Vbs в третий аргумент вставить 0.1, он это будет считывать как true
            StringComparison comparison;
            if (!Convert.ToBoolean(compare))
            {
                comparison = StringComparison.Ordinal;
            }
            else
            {
                comparison = StringComparison.OrdinalIgnoreCase;
            }

            for (int i = 0; i < inputstrings.Length; i++)
            {
                if (inputstrings[i].ToString().IndexOf(value.ToString(), 0, comparison) != -1 && includeToDoulble != 0.0)
                {
                    result.Add(inputstrings[i]);
                }
                else if (inputstrings[i].ToString().IndexOf(value.ToString(), 0, comparison) == -1 && includeToDoulble == 0.0)
                {
                    result.Add(inputstrings[i]);
                }
            }
            return result.ToArray();
        }
        /// <summary>
        ///  Функция Filter возвращает массив из значений отфильтрованных на основе критерии фильтрации.
        /// </summary>
        /// <remarks>
        ///  The Filter function returns a zero-based array that contains a subset of a string array based on a filter criteria.
        /// </remarks>
        /// 
        /// <param name="include">Bool: true - вернет массив, который содержит искомую строку; flase - вернет массив из значений, которые не содержат строку</param>
        /// <returns>Array</returns>
        /// <exception cref="ArgumentException"></exception>
        public static Object[] Filter(object[] inputstrings, object value, object include)
        {
            if (!((value is int || value is long || value is bool || value is double || value is float || value is string)
                && (include is int || include is long || include is bool || include is double || include is float || include is string)))
                throw new ArgumentException();
            List<object> result = new List<object>();
            double includeToDoulble = Convert.ToDouble(include);//Если в Vbs в третий аргумент вставить 0.1, он это будет считывать как true


            for (int i = 0; i < inputstrings.Length; i++)
            {
                if (inputstrings[i].ToString().IndexOf(value.ToString(), 0) != -1 && includeToDoulble != 0.0)
                {
                    result.Add(inputstrings[i]);
                }
                else if (inputstrings[i].ToString().IndexOf(value.ToString(), 0) == -1 && includeToDoulble == 0.0)
                {
                    result.Add(inputstrings[i]);
                }
            }
            return result.ToArray();
        }

        /// <summary>
        ///  Функция Filter возвращает массив из значений отфильтрованных на основе критерии фильтрации.
        /// </summary>
        /// <remarks>
        ///  The Filter function returns a zero-based array that contains a subset of a string array based on a filter criteria.
        /// </remarks>
        /// 
        /// <returns>Array</returns>
        /// <exception cref="ArgumentException"></exception>
        public static Object[] Filter(object[] inputstrings, object value)
        {
            if (!((value is int || value is long || value is bool || value is double || value is float || value is string)))
                throw new ArgumentException();
            List<object> result = new List<object>();

            for (int i = 0; i < inputstrings.Length; i++)
            {
                if (inputstrings[i].ToString().Contains(value.ToString()))
                {
                    result.Add(inputstrings[i]);
                }
                else if (inputstrings[i].ToString().Contains(value.ToString()))
                {
                    result.Add(inputstrings[i]);
                }
            }
            return result.ToArray();
        }

        /// <summary>
        ///  Функция IsArray возвращает true, когда переменная является массивом, и false в противном случае.
        /// </summary>
        /// <remarks>
        ///  The IsArray function returns a Boolean value that indicates whether a specified variable is an array. If the variable is an array, it returns True, otherwise, it returns False.
        /// </remarks>
        /// 
        /// <returns>Boolean</returns>
        public static Boolean IsArray(object obj)
        {
            return obj.GetType().IsArray;
        }

        /// <summary>
        ///  Функция Join возвращает строку, состоящую из элементов массива.
        /// </summary>
        /// <remarks>
        ///  The Join function returns a string that consists of a number of substrings in an array.
        /// </remarks>
        /// <param name="separator">Символы,которые будут отделять элементы</param>
        /// <returns>String</returns>
        public static object Join(object[] array, object separator)
        {
            return String.Join(separator.ToString(), array);
        }
        /// <summary>
        ///  Функция Join возвращает строку, состоящую из элементов массива.
        /// </summary>
        /// <remarks>
        ///  The Join function returns a string that consists of a number of substrings in an array.
        /// </remarks>
        /// <returns>String</returns>
        public static object Join(object[] array)
        {
            return String.Join(" ", array);
        }

        /// <summary>
        ///  Функция LBound возвращает меньшую границу
        /// </summary>
        /// <remarks>
        ///  Returns the smallest subscript for the indicated dimension of an array
        /// </remarks>
        /// <param name="dimension"></param>
        /// <returns>Integer</returns>
        public static object LBound(object array, object dimension)
        {
            if ((!array.GetType().IsArray) || !(dimension is int || dimension is double || dimension is string || dimension is float))
                throw new ArgumentException();
            int dimensionAsInt = Convert.ToInt32(dimension) - 1;
            Array arr = array as Array;
            return arr.GetLowerBound(dimensionAsInt);
        }
        /// <summary>
        ///  Функция LBound возвращает меньшую границу
        /// </summary>
        /// <remarks>
        ///  Returns the smallest subscript for the indicated dimension of an array
        /// </remarks>
        /// <returns>Integer</returns>
        public static object LBound(object array)
        {
            if ((!array.GetType().IsArray))
                throw new ArgumentException();
            Array arr = array as Array;
            return arr.GetLowerBound(0);
        }

        /// <summary>
        /// Функция UBound возвращает большую границу
        /// </summary>
        /// <remarks>
        /// Returns the largest subscript for the indicated dimension of an array
        /// </remarks>
        /// <param name="dimension"></param>
        ///  <returns>Integer</returns>
        public static object UBound(object array, object dimension)
        {
            if ((!array.GetType().IsArray) || !(dimension is int || dimension is double || dimension is string || dimension is float))
                throw new ArgumentException();
            int dimensionAsInt = Convert.ToInt32(dimension) - 1;
            Array arr = array as Array;
            return arr.GetUpperBound(dimensionAsInt);
        }
        /// <summary>
        ///  Функция UBound возвращает большую границу
        /// </summary>
        /// <remarks>
        /// Returns the largest subscript for the indicated dimension of an array
        /// </remarks>
        /// <returns>Integer</returns>
        public static object UBound(object array)
        {
            if ((!array.GetType().IsArray))
                throw new ArgumentException();
            Array arr = array as Array;
            return arr.GetUpperBound(0);
        }

        /// <summary>
        ///  Функция Split возвращает массив, который содержит указанное количество подстрок
        /// </summary>
        /// <remarks>
        /// The Split function returns a zero-based, one-dimensional array that contains a specified number of substrings.
        /// </remarks>
        /// <returns>Array</returns>
        public static object[] Split(object str)
        {
            if (!(str is int || str is long || str is bool || str is double || str is float || str is string))
                throw new ArgumentException();

            return str.ToString().Split(' ');
        }
        /// <summary>
        ///  Функция Split возвращает массив, который содержит указанное количество подстрок
        /// </summary>
        /// <remarks>
        /// The Split function returns a zero-based, one-dimensional array that contains a specified number of substrings.
        /// </remarks>
        /// <returns>Array</returns>
        public static object[] Split(object str, object separator)
        {
            if (!((str is int || str is long || str is bool || str is double || str is float || str is string)
                && (separator is int || separator is long || separator is bool || separator is double || separator is float || separator is string)))
                throw new ArgumentException();
            return str.ToString().Split(Convert.ToChar(separator));
        }
        /// <summary>
        ///  Функция Split возвращает массив, который содержит указанное количество подстрок
        /// </summary>
        /// <remarks>
        /// The Split function returns a zero-based, one-dimensional array that contains a specified number of substrings.
        /// </remarks>
        /// <returns>Array</returns>
        public static object[] Split(object str, object separator,object count)
        {
            if (!((str is int || str is long || str is bool || str is double || str is float || str is string)
                && (separator is int || separator is long || separator is bool || separator is double || separator is float || separator is string)
                && (count is int || count is string || count is float || count is double)))
                throw new ArgumentException();

            char[] separate = separator.ToString().ToCharArray();
            return str.ToString().Split(separate, Convert.ToInt32(count));
        }
        /// <summary>
        ///  Функция Split возвращает массив, который содержит указанное количество подстрок
        /// </summary>
        /// <remarks>
        /// The Split function returns a zero-based, one-dimensional array that contains a specified number of substrings.
        /// </remarks>
        /// <returns>Array</returns>
        public static object[] Split(object str, object separator, object count, object compare)
        {
            if (!((str is int || str is long || str is bool || str is double || str is float || str is string)
                && (separator is int || separator is long || separator is bool || separator is double || separator is float || separator is string)
                && (count is int || count is string || count is float || count is double)
                && (compare is int || compare is string || compare is float || compare is double)))
                throw new ArgumentException();

                int compareToInt = Convert.ToInt32(compare);
            if (compareToInt != 0 && compareToInt != 1)
                throw new ArgumentException();

            if (compareToInt == 0)
            {
                char[] separate = separator.ToString().ToCharArray();
                return str.ToString().Split(separate, Convert.ToInt32(count));
            }
            else
            {
                char[] separate = separator.ToString().ToLower().ToCharArray();
                return str.ToString().ToLower().Split(separate, Convert.ToInt32(count));
            }

        }
    }
}
