namespace Excel.AddressParser //.Position

///单元格的位置，基于1而非0
type Position = 
    { row: int
      column: int }
    
    static member MinPosition = 
        { row = 1
          column = 1 }
    
    static member MaxPosition = 
        { row = Row.maxRows
          column = Column.maxColumns }
    
    ///this位于another的左上方
    member this.isTopLeftOf (another: Position) = 
        let r1, c1 = this.row, this.column
        let r2, c2 = another.row, another.column
        r1 <= r2 && c1 <= c2
    
    member this.isBottomRightOf (another: Position) = 
        let r1, c1 = this.row, this.column
        let r2, c2 = another.row, another.column
        r1 >= r2 && c1 >= c2
    
    member this.move (r, c) = 
        let pos = 
            { row = this.row + r
              column = this.column + c }
        if Position.MinPosition.isTopLeftOf pos && pos.isTopLeftOf Position.MaxPosition then pos
        else failwithf "%A" (this,r,c)
    
    member this.A1Style() = sprintf "%s%d" (Column.address this.column) this.row
    member this.R1C1Style() = sprintf "R%dC%d" this.row this.column
