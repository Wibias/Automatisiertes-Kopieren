using System;

namespace Automatisiertes_Kopieren.Helper
{
    public static class StringHelpers
    {
        public static bool AreNamesSimilar(string name1, string name2, int threshold = 2)
        {
            if (name1 == null || name2 == null)
            {
                throw new ArgumentNullException(nameof(name1), "Both strings must not be null.");
            }
            var distance = LevenshteinDistance(name1, name2);

            return distance <= threshold;
        }


        private static int LevenshteinDistance(string s, string t)
        {
            var n = s.Length;
            var m = t.Length;
            var d = new int[n + 1, m + 1];

            if (n == 0) return m;
            if (m == 0) return n;

            for (var i = 0; i <= n; d[i, 0] = i++)
                for (var j = 0; j <= m; d[0, j] = j++)
                    for (var x = 1; x <= n; x++)
                        for (var y = 1; y <= m; y++)
                        {
                            var cost = t[y - 1] == s[x - 1] ? 0 : 1;
                            d[x, y] = Math.Min(Math.Min(d[x - 1, y] + 1, d[x, y - 1] + 1), d[x - 1, y - 1] + cost);
                        }

            return d[n, m];
        }
    }
}