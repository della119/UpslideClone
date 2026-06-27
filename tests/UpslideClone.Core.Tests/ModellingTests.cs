using System.Linq;
using UpslideClone.Core.Modelling;
using Xunit;

namespace UpslideClone.Core.Tests
{
    public class A1Tests
    {
        [Theory]
        [InlineData("A", 1)]
        [InlineData("Z", 26)]
        [InlineData("AA", 27)]
        [InlineData("AZ", 52)]
        [InlineData("BA", 53)]
        [InlineData("XFD", 16384)]
        public void ColumnRoundTrips(string letters, int index)
        {
            Assert.Equal(index, A1.ColumnToIndex(letters));
            Assert.Equal(letters, A1.IndexToColumn(index));
        }
    }

    public class AutocolorClassifierTests
    {
        [Fact]
        public void Constant_IsInput()
        {
            Assert.Equal(CellColorClass.Input, AutocolorClassifier.Classify(false, "1234"));
        }

        [Fact]
        public void SameSheetFormula_IsFormula()
        {
            Assert.Equal(CellColorClass.Formula, AutocolorClassifier.Classify(true, "=A1+B1*2"));
        }

        [Fact]
        public void CrossSheetReference_IsLink()
        {
            Assert.Equal(CellColorClass.Link, AutocolorClassifier.Classify(true, "=Sheet2!A1"));
        }

        [Fact]
        public void ExternalWorkbookReference_IsLink()
        {
            Assert.Equal(CellColorClass.Link, AutocolorClassifier.Classify(true, "=[Book2.xlsx]Sheet1!A1"));
        }

        [Fact]
        public void Empty_IsEmpty()
        {
            Assert.Equal(CellColorClass.Empty, AutocolorClassifier.Classify(false, "", isEmpty: true));
        }

        [Fact]
        public void DefaultHex_BankerColours()
        {
            Assert.Equal("#0000FF", AutocolorClassifier.DefaultHex(CellColorClass.Input));
            Assert.Equal("#000000", AutocolorClassifier.DefaultHex(CellColorClass.Formula));
            Assert.Equal("#008000", AutocolorClassifier.DefaultHex(CellColorClass.Link));
        }
    }

    public class FormulaTransformTests
    {
        [Fact]
        public void WrapIfError_WrapsBody()
        {
            Assert.Equal("=IFERROR(A1/B1,\"\")", FormulaTransform.WrapIfError("=A1/B1"));
        }

        [Fact]
        public void WrapIfError_CustomReplacement()
        {
            Assert.Equal("=IFERROR(A1/B1,0)", FormulaTransform.WrapIfError("=A1/B1", "0"));
        }

        [Fact]
        public void WrapIfError_DoesNotDoubleWrap()
        {
            Assert.Equal("=IFERROR(A1,\"\")", FormulaTransform.WrapIfError("=IFERROR(A1,\"\")"));
        }

        [Fact]
        public void WrapIfError_WrapsWhenIferrorIsNested()
        {
            // IFERROR is only an argument here, so the whole thing should still wrap.
            Assert.Equal("=IFERROR(A1+IFERROR(B1,0),\"\")",
                FormulaTransform.WrapIfError("=A1+IFERROR(B1,0)"));
        }

        [Fact]
        public void WrapIfError_LeavesConstants()
        {
            Assert.Equal("1234", FormulaTransform.WrapIfError("1234"));
        }
    }

    public class FormulaReferencesTests
    {
        [Fact]
        public void Shift_RelativeRefs_Move()
        {
            Assert.Equal("=B3+C3", FormulaReferences.Shift("=A2+B2", 1, 1));
        }

        [Fact]
        public void Shift_AbsoluteAnchors_StayPut()
        {
            Assert.Equal("=$A$1+C3", FormulaReferences.Shift("=$A$1+B2", 1, 1));
        }

        [Fact]
        public void Shift_MixedAnchors()
        {
            // $A1 : column locked, row relative.  A$1 : row locked, column relative.
            Assert.Equal("=$A3+C$1", FormulaReferences.Shift("=$A2+B$1", 1, 1));
        }

        [Fact]
        public void Shift_DoesNotTouchStringsOrFunctions()
        {
            Assert.Equal("=LOG10(B3)&\"A1\"",
                FormulaReferences.Shift("=LOG10(A2)&\"A1\"", 1, 1));
        }

        [Fact]
        public void Shift_AcrossSheetQualifier_ShiftsCellNotSheet()
        {
            Assert.Equal("=Sheet1!B3", FormulaReferences.Shift("=Sheet1!A2", 1, 1));
        }

        [Fact]
        public void Extract_FindsAllRefs()
        {
            var refs = FormulaReferences.Extract("=A1+$B$2+Sheet2!C3");
            var raws = refs.Select(r => r.Raw).ToList();
            Assert.Contains("A1", raws);
            Assert.Contains("$B$2", raws);
            Assert.Contains("C3", raws);
            Assert.Equal(3, refs.Count);
        }

        [Fact]
        public void Extract_IgnoresStrings()
        {
            var refs = FormulaReferences.Extract("=\"see A1\"&B2");
            Assert.Single(refs);
            Assert.Equal("B2", refs[0].Raw);
        }
    }
}
