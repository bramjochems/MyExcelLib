namespace BJExcelLibTests

open System
open Microsoft.VisualStudio.TestTools.UnitTesting
open BJExcelLib.Finance
open BJExcelLib.Finance.Date


[<TestClass>]
type InterpolationTests() =

    [<TestMethod>]
    member x.``Linear fitting reproduces expected values`` () =
        Assert.AreEqual(true,false) //todo

    [<TestMethod>]
    member x.``Polynomial fitting reproduces expected values at various degrees`` () =
        Assert.AreEqual(true,false) //todo

    [<TestMethod>]
    member x.``Polynomial fitting fails on negative degree`` () =
        Assert.AreEqual(true,false) //todo

    [<TestMethod>]
    member x.``Polynomial fitting fails if number of input pairs <= degree of the polynomial to fit`` () =
        Assert.AreEqual(true,false) //todo

    [<TestMethod>]
    member x.``Natural cubic spline produces expected output values`` () =
        Assert.AreEqual(true,false) //todo

    [<TestMethod>]
    member x.``Clamped cubic spline produces expected output values`` () =
        Assert.AreEqual(true,false) //todo

    [<TestMethod>]
    member x.``Hermite cubic spline produces expected output values`` () =
        Assert.AreEqual(true,false) //todo

    [<TestMethod>]
    member x.``Akima cubic spline produces expected output values`` () =
        Assert.AreEqual(true,false) //todo
