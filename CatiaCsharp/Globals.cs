using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CatiaCsharp
{
    public class Globals
    {
        public static int errorCount = 0;
        public enum CATPasteType
        {
            CATPrtCont, //"As Specified In Part Document"
            CATPrtResultWithOutLink, //As Result
            CATPrtResult //As result with link
        }
        public enum Axis
        {
            X,
            Y,
            Z
        }
        public const int PieceLengthInt = 50; //mm
        public static readonly double PieceLengthDouble = (double)PieceLengthInt;
        public const int LengthXPieces = 15;
        public const int LengthYPieces = 10;
        public readonly double BoardXLength = LengthXPieces * PieceLengthDouble;
        public readonly double BoardYWidth = LengthYPieces * PieceLengthDouble;
    }
}
