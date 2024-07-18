using System;
using System.Collections.Generic;
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
        public static object[] Array(params object[] values)
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
        public static object[] Filter(object[] inputstrings, object value, object include, object compare)
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
        public static object[] Filter(object[] inputstrings, object value, object include)
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
        public static object[] Filter(object[] inputstrings, object value)
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


        public static Boolean IsArray(object obj) 
        {
            return obj.GetType().IsArray;
        }

    }
}
