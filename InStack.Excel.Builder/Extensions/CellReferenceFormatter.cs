namespace InStack.Excel.Builder.Extensions;

public static class CellReferenceFormatter
{
    public static int Format(Span<byte> buffer, uint row, uint column)
    {
        unchecked
        {
            column--;

            int letterCount = 1;

            if (column < 26)
            {
                // Single letter: A–Z
                buffer[0] = (byte)(65 + column);
            }
            else if (column < 702)
            {
                // Two letters: AA–ZZ
                uint c1 = column % 26;
                uint c0 = column / 26 - 1;

                buffer[0] = (byte)(65 + c0);
                buffer[1] = (byte)(65 + c1);
                letterCount = 2;
            }
            else
            {
                // Three letters: AAA
                uint rem = column;
                uint c2 = rem % 26;
                rem = rem / 26 - 1;
                uint c1 = rem % 26;
                uint c0 = rem / 26;

                buffer[0] = (byte)(65 + c0);
                buffer[1] = (byte)(65 + c1);
                buffer[2] = (byte)(65 + c2);
                letterCount = 3;
            }

            return letterCount + UintByteFormatter.Format(buffer.Slice(letterCount), row);
        }
    }

}
