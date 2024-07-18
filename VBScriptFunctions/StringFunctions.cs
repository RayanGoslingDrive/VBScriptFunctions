using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace VBScriptFunctions
{
    public class StringFunctions
    {
        /// <summary>
        ///  Функция InStr возвращает позицию первого появления второй строки в первой строке
        /// </summary>
        /// <remarks>
        ///  The InStr function returns the position of the first occurrence of one string within another.
        /// </remarks>
        /// 
        /// If string1 is Null - InStr returns Null, в vbs нет null, только nothing.VBS выдает ошибку "объектная переменная не задана", в C# NullReferenceException 
        /// 
        /// <param name="compare">0 - учет регистра, 1 - без учета регистра</param>
        /// <returns>Integer: возвращаемое значение позиции отсчитывается с начала строки</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object InStr(object start, object string1, object string2, object compare)
        {
            if (!((string1 is int || string1 is long || string1 is bool || string1 is double || string1 is float || string1 is string)
                && (string2 is int || string2 is long || string2 is bool || string2 is double || string2 is float || string2 is string)
                && (start is int || start is string || start is float || start is double)
                && (compare is int || compare is string || compare is float || compare is double)))
                throw new ArgumentException();
            if ((int)compare != 0 && (int)compare != 1)
                throw new ArgumentException();
            if (Convert.ToInt32(start) > string1.ToString().Length)
                return 0;
            if (Convert.ToInt32(compare) == 0)
            {
                int startindex = Convert.ToInt32(start) - 1;
                return string1.ToString().IndexOf(string2.ToString(), startindex, StringComparison.Ordinal) + 1;
            }
            else
            {
                int startindex = Convert.ToInt32(start) - 1;
                return string1.ToString().IndexOf(string2.ToString(), startindex, StringComparison.OrdinalIgnoreCase) + 1;
            }

        }
        /// <summary>
        ///  Функция InStr возвращает позицию первого появления второй строки в первой строке
        /// </summary>
        /// <remarks>
        ///  The InStr function returns the position of the first occurrence of one string within another.
        /// </remarks>
        /// 
        /// If string1 is Null - InStr returns Null, в vbs нет null, только nothing.VBS выдает ошибку "объектная переменная не задана", в C# NullReferenceException 
        /// 
        /// <returns>Integer: возвращаемое значение позиции отсчитывается с начала строки</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object InStr(object start, object string1, object string2)
        {
            if (!((string1 is int || string1 is long || string1 is bool || string1 is double || string1 is float || string1 is string)
                && (string2 is int || string2 is long || string2 is bool || string2 is double || string2 is float || string2 is string)
                && (start is int || start is string || start is float || start is double)))
                throw new ArgumentException();
            if (Convert.ToInt32(start) > string1.ToString().Length)
                return 0;

            int startindex = Convert.ToInt32(start) - 1;
            return string1.ToString().IndexOf(string2.ToString(), startindex) + 1;
        }
        /// <summary>
        ///  Функция InStr возвращает позицию первого появления второй строки в первой строке
        /// </summary>
        /// <remarks>
        ///  The InStr function returns the position of the first occurrence of one string within another.
        /// </remarks>
        /// 
        /// If string1 is Null - InStr returns Null, в vbs нет null, только nothing.VBS выдает ошибку "объектная переменная не задана", в C# NullReferenceException 
        /// 
        /// <returns>Integer: возвращаемое значение позиции отсчитывается с начала строки</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object InStr(object string1, object string2)
        {
            if (!((string1 is int || string1 is long || string1 is bool || string1 is double || string1 is float || string1 is string)
                && (string2 is int || string2 is long || string2 is bool || string2 is double || string2 is float || string2 is string)))
                throw new ArgumentException();

            return string1.ToString().IndexOf(string2.ToString()) + 1;
        }


        /// <summary>
        ///  Функция InStrRev возвращает позицию последнего появления второй строки в первой строке
        /// </summary>
        /// <remarks>
        ///  The InStrRev function returns the position of the first occurrence of one string within another. The search begins from the end of string.
        /// </remarks>
        /// 
        /// If string1 is Null - InStr returns Null, в vbs нет null, только nothing.VBS выдает ошибку "объектная переменная не задана", в C# NullReferenceException 
        /// 
        /// <param name="compare">0 - учет регистра, 1 - без учета регистра</param>
        /// <returns>Integer: возвращаемое значение позиции отсчитывается с начала строки</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object InStrRev(object string1, object string2, object start, object compare)
        {
            if (!((string1 is int || string1 is long || string1 is bool || string1 is double || string1 is float || string1 is string)
                && (string2 is int || string2 is long || string2 is bool || string2 is double || string2 is float || string2 is string)
                && (start is int || start is string || start is float || start is double)
                && (compare is int || compare is string || compare is float || compare is double)))
                throw new ArgumentException();
            if ((int)compare != 0 && (int)compare != 1)
                throw new ArgumentException();
            if (Convert.ToInt32(start) > string1.ToString().Length)
                return 0;


            if (Convert.ToInt32(compare) == 0)
            {
                int startindex;
                if (Convert.ToInt32(start) == -1)
                {
                    startindex = string1.ToString().Length - 1;
                }
                else
                {
                    startindex = Convert.ToInt32(start) - 1;
                }

                return string1.ToString().LastIndexOf(string2.ToString(), startindex, StringComparison.Ordinal) + 1;
            }
            else
            {
                int startindex;
                if (Convert.ToInt32(start) == -1)
                {
                    startindex = string1.ToString().Length - 1;
                }
                else
                {
                    startindex = Convert.ToInt32(start) - 1;
                }
                return string1.ToString().LastIndexOf(string2.ToString(), startindex, StringComparison.OrdinalIgnoreCase) + 1;
            }

        }

        /// <summary>
        ///  Функция InStrRev возвращает позицию последнего появления второй строки в первой строке
        /// </summary>
        /// <remarks>
        ///  The InStrRev function returns the position of the first occurrence of one string within another. The search begins from the end of string.
        /// </remarks>
        /// 
        /// If string1 is Null - InStr returns Null, в vbs нет null, только nothing.VBS выдает ошибку "объектная переменная не задана", в C# NullReferenceException 
        /// 
        /// <returns>Integer: возвращаемое значение позиции отсчитывается с начала строки</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object InStrRev(object string1, object string2, object start)
        {
            if (!((string1 is int || string1 is long || string1 is bool || string1 is double || string1 is float || string1 is string)
                && (string2 is int || string2 is long || string2 is bool || string2 is double || string2 is float || string2 is string)
                && (start is int || start is string || start is float || start is double)))
                throw new ArgumentException();
            if (Convert.ToInt32(start) > string1.ToString().Length)
                return 0;

            int startindex;
            if (Convert.ToInt32(start) == -1)
            {
                startindex = string1.ToString().Length - 1;
            }
            else
            {
                startindex = Convert.ToInt32(start) - 1;
            }
            return string1.ToString().LastIndexOf(string2.ToString(), startindex) + 1;
        }

        /// <summary>
        ///  Функция InStrRev возвращает позицию последнего появления второй строки в первой строке
        /// </summary>
        /// <remarks>
        ///  The InStrRev function returns the position of the first occurrence of one string within another. The search begins from the end of string.
        /// </remarks>
        /// 
        /// If string1 is Null - InStr returns Null, в vbs нет null, только nothing.VBS выдает ошибку "объектная переменная не задана", в C# NullReferenceException 
        /// 
        /// <returns>Integer: возвращаемое значение позиции отсчитывается с начала строки</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object InStrRev(object string1, object string2)
        {
            if (!((string1 is int || string1 is long || string1 is bool || string1 is double || string1 is float || string1 is string) &&
                (string2 is int || string2 is long || string2 is bool || string2 is double || string2 is float || string2 is string)))
                throw new ArgumentException();

            return string1.ToString().LastIndexOf(string2.ToString()) + 1;
        }



        /// <summary>
        ///   Функция LCase конвертирует строку в нижний регистр
        /// </summary>
        /// <remarks>
        ///  The LCase function converts a specified string to lowercase.
        /// </remarks>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object LCase(object str)
        {

            if (!(str is int || str is long || str is bool || str is double || str is float || str is string))
                throw new ArgumentException();

            return str.ToString().ToLower();

        }




        /// <summary>
        ///   Функция Left возвращает указанное количество символов в строке, начиная с левой стороны
        /// </summary>
        /// <remarks>
        ///  The Left function returns a specified number of characters from the left side of a string.
        /// </remarks>
        /// <param name="length">длина возвращаемой строки.</param>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object Left(object str, object length)
        {


            if (!((str is int || str is long || str is bool || str is double || str is float || str is string)
                && (length is int || length is string || length is float || length is double)))
                throw new ArgumentException();

            if (Convert.ToInt32(length) >= str.ToString().Length)
                return str;

            return str.ToString().Remove(Convert.ToInt32(length));

        }

        /// <summary>
        ///   Функция Right возвращает указанное количество символов в строке, начиная с правой стороны
        /// </summary>
        /// <remarks>
        /// The Right function returns a specified number of characters from the right side of a string.
        /// </remarks>
        /// <param name="length">длина возвращаемой строки.</param>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object Right(object str, object length)
        {

            if (!((str is int || str is long || str is bool || str is double || str is float || str is string)
                && (length is int || length is string || length is float || length is double)))
                throw new ArgumentException();

            if (Convert.ToInt32(length) >= str.ToString().Length)
                return str;

            return str.ToString().Remove(0, str.ToString().Length - Convert.ToInt32(length));

        }



        /// <summary>
        ///  Функция Len возвращает количество символов в строке
        /// </summary>
        /// <remarks>
        /// The Len function returns the number of characters in a string.
        /// </remarks>
        /// <returns>Integer</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object Len(object str)
        {

            if (!(str is int || str is long || str is bool || str is double || str is float || str is string))
                throw new ArgumentException();
            return str.ToString().Length;
        }


        /// <summary>
        ///   Функция LTrim убирает пробелы в левой части строчки
        /// </summary>
        /// <remarks>
        /// The LTrim function removes spaces on the left side of a string.
        /// </remarks>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object LTrim(object str)
        {

            if (!(str is int || str is long || str is bool || str is double || str is float || str is string))
                throw new ArgumentException();
            return str.ToString().TrimStart(' ');
        }


        /// <summary>
        ///   Функция RTrim убирает пробелы в правой части строчки
        /// </summary>
        /// <remarks>
        /// The RTrim function removes spaces on the right side of a string.
        /// </remarks>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object RTrim(object str)
        {

            if (!(str is int || str is long || str is bool || str is double || str is float || str is string))
                throw new ArgumentException();
            return str.ToString().TrimEnd(' ');
        }


        /// <summary>
        ///   Функция Trim убирает пробелы с обеих сторон строчки
        /// </summary>
        /// <remarks>
        /// The Trim function removes spaces on both sides of a string.
        /// </remarks>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object Trim(object str)
        {

            if (!(str is int || str is long || str is bool || str is double || str is float || str is string))
                throw new ArgumentException();
            return str.ToString().TrimStart(' ').TrimEnd(' ');
        }


        /// <summary>
        ///   Функция Mid возвращает заданное количество символов из строки
        /// </summary>
        /// <remarks>
        /// The Mid function returns a specified number of characters from a string.
        /// </remarks>
        /// <param name="length">длина возвращаемой строки</param>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object Mid(object str, object start, object length)
        {

            if (!((str is int || str is long || str is bool || str is double || str is float || str is string)
             && (start is int || start is string || start is float || start is double)
             && (length is int || length is string || length is float || length is double)))
                throw new ArgumentException();
            if (Convert.ToInt32(start) >= str.ToString().Length)
                return "";
            if (Convert.ToInt32(length) > str.ToString().Length)
                return str;

            return str.ToString().Substring(Convert.ToInt32(start) - 1, Convert.ToInt32(length));
        }
        /// <summary>
        ///   Функция Mid возвращает заданное количество символов из строки
        /// </summary>
        /// <remarks>
        /// The Mid function returns a specified number of characters from a string.
        /// </remarks>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object Mid(object str, object start)
        {

            if (!((str is int || str is long || str is bool || str is double || str is float || str is string)
                && (start is int || start is string || start is float || start is double)))
                throw new ArgumentException();
            if (Convert.ToInt32(start) >= str.ToString().Length)
                return "";

            return str.ToString().Substring(Convert.ToInt32(start) - 1);
        }


        /// <summary>
        ///  Функция Replace заменяет часть строки на другую строку, указанное количество раз
        /// </summary>
        /// <remarks>
        /// The Replace function replaces a specified part of a string with another string a specified number of times.
        /// </remarks>
        /// <param name="str"></param>
        /// <param name="compare">0 - учет регистра, 1 - без учета регистра</param>
        /// <param name="count">Сколько найденных строк нужно заменить</param>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object Replace(object str, object find, object replacewith, object start, object count, object compare)
        {


            if (!((str is int || str is long || str is bool || str is double || str is float || str is string)
                && (find is int || find is long || find is bool || find is double || find is float || find is string)
                && (replacewith is int || replacewith is long || replacewith is bool || replacewith is double || replacewith is float || replacewith is string)
                && (start is int || start is string || start is float || start is double)
                && (count is int || count is string || count is float || count is double)
                && (compare is int || compare is string || compare is float || compare is double)))
                throw new ArgumentException();
            if (Convert.ToInt32(compare) != 0 && Convert.ToInt32(compare) != 1)
                throw new ArgumentException();
            if (Convert.ToInt32(start) >= str.ToString().Length)
                return "";

            string result = "";
            Regex regex;
            bool comparator = Convert.ToBoolean(compare);
            if (!comparator)
            {
                regex = new Regex(find.ToString());
                result = regex.Replace(str.ToString().Substring(Convert.ToInt32(start) - 1), replacewith.ToString(), (int)count);
            }
            else
            {
                regex = new Regex(find.ToString().ToLower());
                result = regex.Replace(str.ToString().Substring(Convert.ToInt32(start) - 1).ToLower(), replacewith.ToString(), (int)count);
            }

            return result;
        }

        /// <summary>
        ///  Функция Replace заменяет часть строки на другую строку, указанное количество раз
        /// </summary>
        /// <remarks>
        /// The Replace function replaces a specified part of a string with another string a specified number of times.
        /// </remarks>
        /// <param name="str"></param>
        /// <param name="find"></param>
        /// <param name="count">Сколько найденных строк нужно заменить</param>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object Replace(object str, object find, object replacewith, object start, object count)
        {

            if (!((str is int || str is long || str is bool || str is double || str is float || str is string)
                && (find is int || find is long || find is bool || find is double || find is float || find is string)
                && (replacewith is int || replacewith is long || replacewith is bool || replacewith is double || replacewith is float || replacewith is string)
                && (start is int || start is string || start is float || start is double)
                && (count is int || count is string || count is float || count is double)))
                throw new ArgumentException();
            if (Convert.ToInt32(start) >= str.ToString().Length)
                return "";
            string result = "";
            Regex regex;

            regex = new Regex(find.ToString());
            result = regex.Replace(str.ToString().Substring(Convert.ToInt32(start) - 1), replacewith.ToString(), (int)count);

            return result;
        }


        /// <summary>
        ///  Функция Replace заменяет часть строки на другую строку, указанное количество раз
        /// </summary>
        /// <remarks>
        /// The Replace function replaces a specified part of a string with another string a specified number of times.
        /// </remarks>
        /// <param name="str"></param>
        /// <param name="find"></param>
        /// <param name="replacewith"></param>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object Replace(object str, object find, object replacewith, object start)
        {

            if (!((str is int || str is long || str is bool || str is double || str is float || str is string)
                && (find is int || find is long || find is bool || find is double || find is float || find is string)
                && (replacewith is int || replacewith is long || replacewith is bool || replacewith is double || replacewith is float || replacewith is string)
                && (start is int || start is string || start is float || start is double)))
                throw new ArgumentException();
            if (Convert.ToInt32(start) >= str.ToString().Length)
                return "";

            return str.ToString().Substring(Convert.ToInt32(start) - 1).Replace(find.ToString(), replacewith.ToString());
        }

        /// <summary>
        ///  Функция Replace заменяет часть строки на другую строку, указанное количество раз
        /// </summary>
        /// <remarks>
        /// The Replace function replaces a specified part of a string with another string a specified number of times.
        /// </remarks>
        /// <param name="str"></param>
        /// <param name="find"></param>
        /// <param name="replacewith"></param>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object Replace(object str, object find, object replacewith)
        {

            if (!((str is int || str is long || str is bool || str is double || str is float || str is string)
                && (find is int || find is long || find is bool || find is double || find is float || find is string)
                && (replacewith is int || replacewith is long || replacewith is bool || replacewith is double || replacewith is float || replacewith is string)))
                throw new ArgumentException();


            return str.ToString().Replace(find.ToString(), replacewith.ToString());
        }



        /// <summary>
        /// Функция Space возвращает строку из указанного количества пробелов
        /// </summary>
        /// <remarks>
        /// The Space function returns a string that consists of a specified number of spaces.
        /// </remarks>
        /// <param name="number"></param>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object Space(object number)
        {
            if (!(number is int || number is double || number is float || number is string))
                throw new ArgumentException();
            object result = new string(' ', Convert.ToInt32(number));
            return result;
        }

        /// <summary>
        /// Функция StrComp сравнивает 2 строки и возвращает значение, отражающее результат сравнения
        /// </summary>
        /// <remarks>
        /// The StrComp function compares two strings and returns a value that represents the result of the comparison.
        /// </remarks>
        /// 
        /// If string1 or string2 is Null - InStr returns Null, в vbs нет null, только nothing.VBS выдает ошибку "объектная переменная не задана", в C# NullReferenceException 
        /// 
        /// <param name="string2"></param>
        /// <param name="string2"></param>
        /// <param name="compare">0 - учет регистра, 1 - без учета регистра</param>
        /// <returns>Integer : 0 если строки равны; -1 если вторая строка больше первой; 1 если первая строка больше второй</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object StrComp(object string1, object string2, object compare)
        {
            if (!((string1 is int || string1 is long || string1 is bool || string1 is double || string1 is float || string1 is string)
                && (string2 is int || string2 is long || string2 is bool || string2 is double || string2 is float || string2 is string)
                && (compare is int || compare is string || compare is float || compare is double)))
                throw new ArgumentException();
            int compareToInt = Convert.ToInt32(compare);
            if (compareToInt != 0 && compareToInt != 1)
                throw new ArgumentException();
            if (compareToInt == 0)
            {
                return CultureInfo.CurrentCulture.CompareInfo.Compare(string2.ToString(), string1.ToString(), CompareOptions.None);
            }
            else
            {
                return CultureInfo.CurrentCulture.CompareInfo.Compare(string2.ToString(), string1.ToString(), CompareOptions.IgnoreCase);
            }

        }

        /// <summary>
        /// Функция StrComp сравнивает 2 строки и возвращает значение, отражающее результат сравнения
        /// </summary>
        /// <remarks>
        /// The StrComp function compares two strings and returns a value that represents the result of the comparison.
        /// </remarks>
        /// 
        /// If string1 or string2 is Null - InStr returns Null, в vbs нет null, только nothing.VBS выдает ошибку "объектная переменная не задана", в C# NullReferenceException 
        /// 
        /// <param name="string1"></param>
        /// <param name="string2"></param>
        /// <returns>Integer : 0 если строки равны; -1 если вторая строка больше первой; 1 если первая строка больше второй</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object StrComp(object string1, object string2)
        {
            if (!((string1 is int || string1 is long || string1 is bool || string1 is double || string1 is float || string1 is string)
                && (string2 is int || string2 is long || string2 is bool || string2 is double || string2 is float || string2 is string)))
                throw new ArgumentException();

            return CultureInfo.CurrentCulture.CompareInfo.Compare(string2.ToString(), string1.ToString(), CompareOptions.None);
        }


        /// <summary>
        /// Функция UCase конвертирует строку в верхний регистр
        /// </summary>
        /// <remarks>
        /// The UCase function converts a specified string to uppercase.
        /// </remarks>
        /// <param name="number"></param>
        /// <param name="character"></param>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object StrReverse(object str)
        {
            if (!(str is int || str is long || str is bool || str is double || str is float || str is string))
                throw new ArgumentException();
            char[] ans = str.ToString().ToCharArray();

            Array.Reverse(ans);
            return new String(ans);
        }




        /// <summary>
        /// Функция UCase конвертирует строку в верхний регистр
        /// </summary>
        /// <remarks>
        /// The UCase function converts a specified string to uppercase.
        /// </remarks>
        /// <param name="number"></param>
        /// <param name="character"></param>
        /// <returns>string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object UCase(object str)
        {
            if (!(str is int || str is long || str is bool || str is double || str is float || str is string))
                throw new ArgumentException();
            return str.ToString().ToUpper();

        }



        /// <summary>
        /// Функция String возвращает строчку, которая состоит из заданного количеств определенных символов.
        /// </summary>
        /// <remarks>
        /// The String function returns a string that contains a repeating character of a specified length.
        /// </remarks>
        /// <param name="number"></param>
        /// <param name="character"></param>
        /// <returns>Тип string</returns>
        /// <exception cref="ArgumentException"></exception>
        public static object String(object number, object character)
        {
            if (!(character is int || character is long || character is bool || character is double || character is float || character is string))
                throw new ArgumentException();
            if (character is int)// Vbs рассматривает int как char
                return new String(Convert.ToChar((int)character), (int)number);

            return new String(character.ToString()[0], (int)number);

        }
    }
}
