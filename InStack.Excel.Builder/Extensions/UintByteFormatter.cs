using System.Runtime.CompilerServices;

namespace InStack.Excel.Builder.Extensions;

public static class UintByteFormatter
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int Format(Span<byte> buffer, uint val)
    {
        unchecked
        {
            if (val < 9)
            {
                buffer[0] = (byte)(val + 48);

                return 1;
            }
            else if (val < 100)
            {
                buffer[1] = (byte)(val % 10 + 48);
                buffer[0] = (byte)((val /= 10) % 10 + 48);

                return 2;

            }
            else if (val < 1000)
            {
                buffer[2] = (byte)(val % 10 + 48);
                buffer[1] = (byte)((val /= 10) % 10 + 48);
                buffer[0] = (byte)((val /= 10) % 10 + 48);

                return 3;

            }
            else if (val < 10000)
            {
                buffer[3] = (byte)(val % 10 + 48);
                buffer[2] = (byte)((val /= 10) % 10 + 48);
                buffer[1] = (byte)((val /= 10) % 10 + 48);
                buffer[0] = (byte)((val /= 10) % 10 + 48);

                return 4;
            }
            else if (val < 100000)
            {
                buffer[4] = (byte)(val % 10 + 48);
                buffer[3] = (byte)((val /= 10) % 10 + 48);
                buffer[2] = (byte)((val /= 10) % 10 + 48);
                buffer[1] = (byte)((val /= 10) % 10 + 48);
                buffer[0] = (byte)((val /= 10) % 10 + 48);

                return 5;
            }
            else if (val < 1000000)
            {
                buffer[5] = (byte)(val % 10 + 48);
                buffer[4] = (byte)((val /= 10) % 10 + 48);
                buffer[3] = (byte)((val /= 10) % 10 + 48);
                buffer[2] = (byte)((val /= 10) % 10 + 48);
                buffer[1] = (byte)((val /= 10) % 10 + 48);
                buffer[0] = (byte)((val /= 10) % 10 + 48);

                return 6;
            }
            else
            {
                buffer[6] = (byte)(val % 10 + 48);
                buffer[5] = (byte)((val /= 10) % 10 + 48);
                buffer[4] = (byte)((val /= 10) % 10 + 48);
                buffer[3] = (byte)((val /= 10) % 10 + 48);
                buffer[2] = (byte)((val /= 10) % 10 + 48);
                buffer[1] = (byte)((val /= 10) % 10 + 48);
                buffer[0] = (byte)((val /= 10) % 10 + 48);

                return 7;
            }
        }
    }

}
