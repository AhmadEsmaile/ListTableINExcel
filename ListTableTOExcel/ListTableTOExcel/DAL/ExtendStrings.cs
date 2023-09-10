using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListTableTOExcel.DAL
{
    public static class ExtendStrings
    {
        /// <summary>
        /// Determines whether this string is contained in the specified text. 
        /// </summary>
        /// <param name="str">The string.</param>
        /// <param name="container">The container.</param>
        /// <param name="caseSensitive">if set to <c>true</c> [case sensitive].</param>
        /// <returns>True if the string is found in the container</returns>
        public static bool IsIn(this string str, string container, bool caseSensitive = true)
        {
            bool result = !string.IsNullOrEmpty(str);
            if (result)
            {
                result = (caseSensitive) ? (str == container) || container.IndexOf(str) >= 0
                                         : str.ToUpper() == container.ToUpper() || container.ToUpper().IndexOf(str.ToUpper()) >= 0;
            }
            return result;
        }

        /// <summary>
        /// Determines if the string is contained in the specified string collection. 
        /// </summary>
        /// <param name="str">The string</param>
        /// <param name="list">The string collection</param>
        /// <param name="caseSensitive">Flag indicating whether the comparison is case-sensitive </param>
        /// <returns>True if the string is found in the collection</returns>
        public static bool IsIn(this string str, IEnumerable<string> list, bool caseSensitive = true)
        {
            return (caseSensitive) ? (list.Where(x => x == str).FirstOrDefault() != null)
                                   : (list.Where(x => x.ToUpper() == str.ToUpper()).FirstOrDefault() != null);
        }

        /// <summary>
        /// Encode the specified text to base64. 
        /// </summary>
        /// <param name="text">The text to encode</param>
        /// <param name="encoding">What encoding to use for the text</param>
        /// <returns>If the text is null/empty, an empty string, otherwise the base4 encoded string.</returns>
        public static string Base64Encode(this string text, Encoding encoding = null)
        {
            string value = string.Empty;
            if (!string.IsNullOrEmpty(text))
            {
                encoding = (encoding == null) ? Encoding.UTF8 : encoding;
                byte[] bytes = encoding.GetBytes(text);
                value = Convert.ToBase64String(bytes);
            }
            return value;
        }

        /// <summary>
        /// Decode the specified text from a base64 value
        /// </summary>
        /// <param name="text">The text to encode<</param>
        /// <param name="encoding">What encoding to use for the text<</param>
        /// <returns>If the text is not a valid base64 string, the original text is returned. Otherwise the base4 decoded string is returned.</returns>
        public static string Base64Decode(this string text, Encoding encoding = null)
        {
            string value = string.Empty;
            byte[] bytes;
            if (!string.IsNullOrEmpty(text))
            {
                encoding = (encoding == null) ? Encoding.UTF8 : encoding;
                try
                {
                    bytes = Convert.FromBase64String(text);
                    value = encoding.GetString(bytes);
                }
                catch (Exception)
                {
                    value = text;
                }
            }
            return value;
        }

        /// <summary>
        /// Determines whether this string is "like" the specified string. Performs 
        /// a SQL "LIKE" comparison. 
        /// </summary>
        /// <param name="str">This string.</param>
        /// <param name="like">The string to compare it against.</param>
        /// <returns></returns>
        public static bool IsLike(this string str, string pattern)
        {
            // this code is much faster than a regular expression that performs the same comparison.
            bool isMatch = true;
            bool isWildCardOn = false;
            bool isCharWildCardOn = false;
            bool isCharSetOn = false;
            bool isNotCharSetOn = false;
            bool endOfPattern = false;
            int lastWildCard = -1;
            int patternIndex = 0;
            char p = '\0';
            List<char> set = new List<char>();

            for (int i = 0; i < str.Length; i++)
            {
                char c = str[i];
                endOfPattern = (patternIndex >= pattern.Length);
                if (!endOfPattern)
                {
                    p = pattern[patternIndex];
                    if (!isWildCardOn && p == '%')
                    {
                        lastWildCard = patternIndex;
                        isWildCardOn = true;
                        while (patternIndex < pattern.Length && pattern[patternIndex] == '%')
                        {
                            patternIndex++;
                        }
                        p = (patternIndex >= pattern.Length) ? '\0' : p = pattern[patternIndex];
                    }
                    else if (p == '_')
                    {
                        isCharWildCardOn = true;
                        patternIndex++;
                    }
                    else if (p == '[')
                    {
                        if (pattern[++patternIndex] == '^')
                        {
                            isNotCharSetOn = true;
                            patternIndex++;
                        }
                        else
                        {
                            isCharSetOn = true;
                        }

                        set.Clear();
                        if (pattern[patternIndex + 1] == '-' && pattern[patternIndex + 3] == ']')
                        {
                            char start = char.ToUpper(pattern[patternIndex]);
                            patternIndex += 2;
                            char end = char.ToUpper(pattern[patternIndex]);
                            if (start <= end)
                            {
                                for (char ci = start; ci <= end; ci++)
                                {
                                    set.Add(ci);
                                }
                            }
                            patternIndex++;
                        }

                        while (patternIndex < pattern.Length && pattern[patternIndex] != ']')
                        {
                            set.Add(pattern[patternIndex]);
                            patternIndex++;
                        }
                        patternIndex++;
                    }
                }

                if (isWildCardOn)
                {
                    if (char.ToUpper(c) == char.ToUpper(p))
                    {
                        isWildCardOn = false;
                        patternIndex++;
                    }
                }
                else if (isCharWildCardOn)
                {
                    isCharWildCardOn = false;
                }
                else if (isCharSetOn || isNotCharSetOn)
                {
                    bool charMatch = (set.Contains(char.ToUpper(c)));
                    if ((isNotCharSetOn && charMatch) || (isCharSetOn && !charMatch))
                    {
                        if (lastWildCard >= 0)
                        {
                            patternIndex = lastWildCard;
                        }
                        else
                        {
                            isMatch = false;
                            break;
                        }
                    }
                    isNotCharSetOn = isCharSetOn = false;
                }
                else
                {
                    if (char.ToUpper(c) == char.ToUpper(p))
                    {
                        patternIndex++;
                    }
                    else
                    {
                        if (lastWildCard >= 0)
                        {
                            patternIndex = lastWildCard;
                        }
                        else
                        {
                            isMatch = false;
                            break;
                        }
                    }
                }
            }
            endOfPattern = (patternIndex >= pattern.Length);

            if (isMatch && !endOfPattern)
            {
                bool isOnlyWildCards = true;
                for (int i = patternIndex; i < pattern.Length; i++)
                {
                    if (pattern[i] != '%')
                    {
                        isOnlyWildCards = false;
                        break;
                    }
                }
                if (isOnlyWildCards)
                {
                    endOfPattern = true;
                }
            }
            return (isMatch && endOfPattern);
        }

        /// <summary>
        /// Finds the index of the first numeric character.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns></returns>
        public static int IndexOfFirstNumber(this string str)
        {
            //int index = -1;
            //string numbers = "0123456789";
            //foreach(char ch in str)
            //{
            //	if (ch.ToString().IsIn(numbers))
            //	{
            //		index = str.IndexOf(ch);
            //	}
            //}
            string numbers = "0123456789";
            int index = str.IndexOfAny(numbers.ToArray());
            return index;
        }

        /// <summary>
        /// Finds the index of the first alphabetic character.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns></returns>
        public static int IndexOfFirstAlpha(this string str, bool caseSensitive = false)
        {
            //int index = -1;
            //string alphas = "abcdefghijklmnopqrstuvwxyz";
            //foreach(char ch in str)
            //{
            //	if (ch.ToString().IsIn(alphas))
            //	{
            //		index = str.IndexOf(ch);
            //	}
            //}
            string alphas = "abcdefghijklmnopqrstuvwxyz";
            int index = str.IndexOfAny(alphas.ToArray());
            return index;
        }

        /// <summary>
        /// Finds the index of the first special character.
        /// </summary>
        /// <param name="str">The string.</param>
        /// <returns></returns>
        public static int IndexOfFirstSpecial(this string str)
        {
            string alphas = @"~`!@#$%^&*()_+-=[]\{}|;':"",./<>?";
            int index = str.IndexOfAny(alphas.ToArray());
            return index;
        }

    }
}
