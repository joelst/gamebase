namespace BAC
{
    internal static class Properties
    {
        internal static class Resources
        {
            private static System.Drawing.Bitmap? _master;
            private static System.Drawing.Bitmap? _diff;

            /// <summary>
            /// Gets the Master toolbar icon.
            /// </summary>
            public static System.Drawing.Bitmap master
            {
                get
                {
                    if (_master == null)
                    {
                        // Create a simple placeholder icon (32x32, blue)
                        _master = new System.Drawing.Bitmap(32, 32);
                        using (var g = System.Drawing.Graphics.FromImage(_master))
                        {
                            g.Clear(System.Drawing.Color.Blue);
                            g.DrawString("M", new System.Drawing.Font("Arial", 16), System.Drawing.Brushes.White, 4, 4);
                        }
                    }
                    return _master;
                }
            }

            /// <summary>
            /// Gets the Diff toolbar icon.
            /// </summary>
            public static System.Drawing.Bitmap diff
            {
                get
                {
                    if (_diff == null)
                    {
                        // Create a simple placeholder icon (32x32, green)
                        _diff = new System.Drawing.Bitmap(32, 32);
                        using (var g = System.Drawing.Graphics.FromImage(_diff))
                        {
                            g.Clear(System.Drawing.Color.Green);
                            g.DrawString("D", new System.Drawing.Font("Arial", 16), System.Drawing.Brushes.White, 4, 4);
                        }
                    }
                    return _diff;
                }
            }
        }
    }
}
