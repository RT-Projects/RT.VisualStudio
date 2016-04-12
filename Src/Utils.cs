using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RT.Util
{
    /// <summary>This class offers some generic static functions which are hard to categorize under any more specific classes.</summary>
    static class Ut
    {
        /// <summary>
        ///     Returns the parameters as a new array.</summary>
        /// <remarks>
        ///     Useful to circumvent Visual Studio’s bug where multi-line literal arrays are not auto-formatted.</remarks>
        public static T[] NewArray<T>(params T[] parameters) { return parameters; }

        /// <summary>
        ///     Throws the specified exception.</summary>
        /// <typeparam name="TResult">
        ///     The type to return.</typeparam>
        /// <param name="exception">
        ///     The exception to throw.</param>
        /// <returns>
        ///     This method never returns a value. It always throws.</returns>
        [DebuggerHidden]
        public static TResult Throw<TResult>(Exception exception)
        {
            throw exception;
        }
    }
}

namespace RT.Util.ExtensionMethods
{
    static class IEnumerableExtensions
    {
        /// <summary>
        ///     Accumulates consecutive elements that are equal when processed by a selector.</summary>
        /// <typeparam name="TItem">
        ///     The type of items in the input sequence.</typeparam>
        /// <typeparam name="TKey">
        ///     The return type of the <paramref name="selector"/> function.</typeparam>
        /// <param name="source">
        ///     The input sequence from which to accumulate groups of consecutive elements.</param>
        /// <param name="selector">
        ///     A function to transform each item into a key which is compared for equality.</param>
        /// <param name="keyComparer">
        ///     An optional equality comparer for the keys returned by <paramref name="selector"/>.</param>
        /// <returns>
        ///     A collection containing each sequence of consecutive equal elements.</returns>
        public static IEnumerable<ConsecutiveGroup<TItem, TKey>> GroupConsecutiveBy<TItem, TKey>(this IEnumerable<TItem> source, Func<TItem, TKey> selector, IEqualityComparer<TKey> keyComparer = null)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            var comparer = keyComparer ?? EqualityComparer<TKey>.Default;
            return groupConsecutiveIterator(source, selector, null, keyComparer);
        }

        private static IEnumerable<ConsecutiveGroup<TItem, TKey>> groupConsecutiveIterator<TItem, TKey>(IEnumerable<TItem> source, Func<TItem, TKey> selector, Func<TKey, TKey, bool> itemEquality, IEqualityComparer<TKey> itemComparer)
        {
            bool any = false;
            TKey prevKey = default(TKey);
            var index = 0;
            var currentList = new List<TItem>();
            foreach (var elem in source)
            {
                var key = selector(elem);
                if (!any)
                    any = true;
                else if (itemEquality != null ? !itemEquality(prevKey, key) : itemComparer != null ? !itemComparer.Equals(prevKey, key) : !object.Equals(prevKey, key))
                {
                    yield return new ConsecutiveGroup<TItem, TKey>(index - currentList.Count, currentList, prevKey);
                    currentList = new List<TItem>();
                }
                currentList.Add(elem);
                prevKey = key;
                index++;
            }
            if (any)
                yield return new ConsecutiveGroup<TItem, TKey>(index - currentList.Count, currentList, prevKey);
        }

        /// <summary>
        ///     Turns all elements in the enumerable to strings and joins them using the specified <paramref
        ///     name="separator"/> and the specified <paramref name="prefix"/> and <paramref name="suffix"/> for each string.</summary>
        /// <param name="values">
        ///     The sequence of elements to join into a string.</param>
        /// <param name="separator">
        ///     Optionally, a separator to insert between each element and the next.</param>
        /// <param name="prefix">
        ///     Optionally, a string to insert in front of each element.</param>
        /// <param name="suffix">
        ///     Optionally, a string to insert after each element.</param>
        /// <param name="lastSeparator">
        ///     Optionally, a separator to use between the second-to-last and the last element.</param>
        /// <example>
        ///     <code>
        ///         // Returns "[Paris], [London], [Tokyo]"
        ///         (new[] { "Paris", "London", "Tokyo" }).JoinString(", ", "[", "]")
        ///         
        ///         // Returns "[Paris], [London] and [Tokyo]"
        ///         (new[] { "Paris", "London", "Tokyo" }).JoinString(", ", "[", "]", " and ");</code></example>
        public static string JoinString<T>(this IEnumerable<T> values, string separator = null, string prefix = null, string suffix = null, string lastSeparator = null)
        {
            if (values == null)
                throw new ArgumentNullException("values");
            if (lastSeparator == null)
                lastSeparator = separator;

            using (var enumerator = values.GetEnumerator())
            {
                if (!enumerator.MoveNext())
                    return "";

                // Optimise the case where there is only one element
                var one = enumerator.Current;
                if (!enumerator.MoveNext())
                    return prefix + one + suffix;

                // Optimise the case where there are only two elements
                var two = enumerator.Current;
                if (!enumerator.MoveNext())
                {
                    // Optimise the (common) case where there is no prefix/suffix; this prevents an array allocation when calling string.Concat()
                    if (prefix == null && suffix == null)
                        return one + lastSeparator + two;
                    return prefix + one + suffix + lastSeparator + prefix + two + suffix;
                }

                StringBuilder sb = new StringBuilder()
                    .Append(prefix).Append(one).Append(suffix).Append(separator)
                    .Append(prefix).Append(two).Append(suffix);
                var prev = enumerator.Current;
                while (enumerator.MoveNext())
                {
                    sb.Append(separator).Append(prefix).Append(prev).Append(suffix);
                    prev = enumerator.Current;
                }
                sb.Append(lastSeparator).Append(prefix).Append(prev).Append(suffix);
                return sb.ToString();
            }
        }
    }

    /// <summary>
    ///     Encapsulates information about a group generated by <see
    ///     cref="IEnumerableExtensions.GroupConsecutive{TItem}(IEnumerable{TItem})"/> and its overloads.</summary>
    /// <typeparam name="TItem">
    ///     Type of the elements in the sequence.</typeparam>
    /// <typeparam name="TKey">
    ///     Type of the key by which elements were compared.</typeparam>
    sealed class ConsecutiveGroup<TItem, TKey> : IEnumerable<TItem>
    {
        /// <summary>Index in the original sequence where the group started.</summary>
        public int Index { get; private set; }
        /// <summary>Size of the group.</summary>
        public int Count { get; private set; }
        /// <summary>The key by which the items in this group are deemed equal.</summary>
        public TKey Key { get; private set; }

        private IEnumerable<TItem> _group;
        internal ConsecutiveGroup(int index, List<TItem> group, TKey key)
        {
            Index = index;
            Count = group.Count;
            Key = key;
            _group = group;
        }

        /// <summary>
        ///     Returns an enumerator that iterates through the collection.</summary>
        /// <returns>
        ///     A <see cref="System.Collections.Generic.IEnumerator{T}"/> that can be used to iterate through the collection.</returns>
        public IEnumerator<TItem> GetEnumerator() { return _group.GetEnumerator(); }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { return GetEnumerator(); }

        /// <summary>Returns a string that represents this group’s key and its count.</summary>
        public override string ToString()
        {
            return "{0}; Count = {1}".Fmt(Key, Count);
        }
    }

    static class StringExtensions
    {
        /// <summary>Formats a string in a way compatible with <see cref="string.Format(string, object[])"/>.</summary>
        public static string Fmt(this string formatString, params object[] args)
        {
            return string.Format(formatString, args);
        }

        /// <summary>
        ///     Returns a string array that contains the substrings in this string that are delimited by elements of a
        ///     specified string array.</summary>
        /// <param name="str">
        ///     String to be split.</param>
        /// <param name="separator">
        ///     Strings that delimit the substrings in this string.</param>
        /// <returns>
        ///     An array whose elements contain the substrings in this string that are delimited by one or more strings in
        ///     separator. For more information, see the Remarks section.</returns>
        public static string[] Split(this string str, params string[] separator)
        {
            return str.Split(separator, StringSplitOptions.None);
        }

        /// <summary>
        ///     Word-wraps the current string to a specified width. Supports UNIX-style newlines and indented paragraphs.</summary>
        /// <remarks>
        ///     <para>
        ///         The supplied text will be split into "paragraphs" at the newline characters. Every paragraph will begin on
        ///         a new line in the word-wrapped output, indented by the same number of spaces as in the input. All
        ///         subsequent lines belonging to that paragraph will also be indented by the same amount.</para>
        ///     <para>
        ///         All multiple contiguous spaces will be replaced with a single space (except for the indentation).</para></remarks>
        /// <param name="text">
        ///     Text to be word-wrapped.</param>
        /// <param name="maxWidth">
        ///     The maximum number of characters permitted on a single line, not counting the end-of-line terminator.</param>
        /// <param name="hangingIndent">
        ///     The number of spaces to add to each line except the first of each paragraph, thus creating a hanging
        ///     indentation.</param>
        public static IEnumerable<string> WordWrap(this string text, int maxWidth, int hangingIndent = 0)
        {
            if (text == null)
                throw new ArgumentNullException("text");
            if (maxWidth < 1)
                throw new ArgumentOutOfRangeException("maxWidth", maxWidth, "maxWidth cannot be less than 1");
            if (hangingIndent < 0)
                throw new ArgumentOutOfRangeException("hangingIndent", hangingIndent, "hangingIndent cannot be negative.");
            if (text == null || text.Length == 0)
                return Enumerable.Empty<string>();

            return wordWrap(
                text.Split(new string[] { "\r\n", "\r", "\n" }, StringSplitOptions.None),
                maxWidth,
                hangingIndent,
                (txt, substrIndex) => txt.Substring(substrIndex).Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries),
                str => str.Length,
                txt =>
                {
                    // Count the number of spaces at the start of the paragraph
                    int indentLen = 0;
                    while (indentLen < txt.Length && txt[indentLen] == ' ')
                        indentLen++;
                    return indentLen;
                },
                num => new string(' ', num),
                () => new StringBuilder(),
                sb => sb.Length,
                (sb, str) => { sb.Append(str); },
                sb => sb.ToString(),
                (str, start, length) => length == null ? str.Substring(start) : str.Substring(start, length.Value),
                (str1, str2) => str1 + str2);
        }

        private static IEnumerable<T> wordWrap<T, TBuilder>(IEnumerable<T> paragraphs, int maxWidth, int hangingIndent, Func<T, int, IEnumerable<T>> splitSubstringIntoWords,
            Func<T, int> getLength, Func<T, int> getIndent, Func<int, T> spaces, Func<TBuilder> getBuilder, Func<TBuilder, int> getTotalLength, Action<TBuilder, T> add,
            Func<TBuilder, T> getString, Func<T, int, int?, T> substring, Func<T, T, T> concat)
        {
            foreach (var paragraph in paragraphs)
            {
                var indentLen = getIndent(paragraph);
                var indent = spaces(indentLen + hangingIndent);
                var space = spaces(indentLen);
                var numSpaces = indentLen;
                var curLine = getBuilder();

                // Split into words
                foreach (var wordForeach in splitSubstringIntoWords(paragraph, indentLen))
                {
                    var word = wordForeach;
                    var curLineLength = getTotalLength(curLine);

                    if (curLineLength + numSpaces + getLength(word) > maxWidth)
                    {
                        // Need to wrap
                        if (getLength(word) > maxWidth)
                        {
                            // This is a very long word
                            // Leave part of the word on the current line if at least 2 chars fit
                            if (curLineLength + numSpaces + 2 <= maxWidth || getTotalLength(curLine) == 0)
                            {
                                int length = maxWidth - getTotalLength(curLine) - numSpaces;
                                add(curLine, space);
                                add(curLine, substring(word, 0, length));
                                word = substring(word, length, null);
                            }
                            // Commit the current line
                            yield return getString(curLine);

                            // Now append full lines' worth of text until we're left with less than a full line
                            while (indentLen + getLength(word) > maxWidth)
                            {
                                yield return concat(indent, substring(word, 0, maxWidth - indentLen));
                                word = substring(word, maxWidth - indentLen, null);
                            }

                            // Start a new line with whatever is left
                            curLine = getBuilder();
                            add(curLine, indent);
                            add(curLine, word);
                        }
                        else
                        {
                            // This word is not very long and it doesn't fit so just wrap it to the next line
                            yield return getString(curLine);

                            // Start a new line
                            curLine = getBuilder();
                            add(curLine, indent);
                            add(curLine, word);
                        }
                    }
                    else
                    {
                        // No need to wrap yet
                        add(curLine, space);
                        add(curLine, word);
                    }

                    if (numSpaces != 1)
                    {
                        space = spaces(1);
                        numSpaces = 1;
                    }
                }

                yield return getString(curLine);
            }
        }

        /// <summary>
        ///     Inserts spaces at the beginning of every line contained within the specified string.</summary>
        /// <param name="str">
        ///     String to add indentation to.</param>
        /// <param name="by">
        ///     Number of spaces to add.</param>
        /// <param name="indentFirstLine">
        ///     If true (default), all lines are indented; otherwise, all lines except the first.</param>
        /// <returns>
        ///     The indented string.</returns>
        public static string Indent(this string str, int by, bool indentFirstLine = true)
        {
            if (indentFirstLine)
                return Regex.Replace(str, "^", new string(' ', by), RegexOptions.Multiline);
            return Regex.Replace(str, "(?<=\n)", new string(' ', by));
        }

        /// <summary>
        ///     Escapes all necessary characters in the specified string so as to make it usable safely in an HTML or XML
        ///     context.</summary>
        /// <param name="input">
        ///     The string to apply HTML or XML escaping to.</param>
        /// <param name="leaveSingleQuotesAlone">
        ///     If <c>true</c>, does not escape single quotes (<c>'</c>, U+0027).</param>
        /// <param name="leaveDoubleQuotesAlone">
        ///     If <c>true</c>, does not escape single quotes (<c>"</c>, U+0022).</param>
        /// <returns>
        ///     The specified string with the necessary HTML or XML escaping applied.</returns>
        public static string HtmlEscape(this string input, bool leaveSingleQuotesAlone = false, bool leaveDoubleQuotesAlone = false)
        {
            if (input == null)
                throw new ArgumentNullException("input");
            var result = input.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
            if (!leaveSingleQuotesAlone)
                result = result.Replace("'", "&#39;");
            if (!leaveDoubleQuotesAlone)
                result = result.Replace("\"", "&quot;");
            return result;
        }
    }
}
