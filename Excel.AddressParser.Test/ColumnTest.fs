namespace Excel.AddressParser.Test
open Excel.AddressParser

open Xunit
open Xunit.Abstractions

type ColumnTest(output : ITestOutputHelper) =
    [<Fact>]
    member this.``test valueof``() =
        Assert.Equal(Column.valueof 'A',1)
        Assert.Equal(Column.valueof 'a',1)
        Assert.Equal(Column.valueof 'Z',26)
        Assert.Equal(Column.valueof 'z',26)

    [<Fact>]
    member this.``test position``() =
        Assert.Equal(Column.position "A",1)
        Assert.Equal(Column.position "AA",27)
        Assert.Equal(Column.position "AAA",703)
        Assert.Equal(Column.position "XFD",16384)
        Assert.Equal(Column.position "AAAA",18279)

    [<Fact>]
    member this.``test address``() =
        Assert.Equal(Column.address 1    ,"A"   )
        Assert.Equal(Column.address 27   ,"AA"  )
        Assert.Equal(Column.address 703  ,"AAA" )
        Assert.Equal(Column.address 16384,"XFD" )
        Assert.Equal(Column.address 18279,"AAAA")


