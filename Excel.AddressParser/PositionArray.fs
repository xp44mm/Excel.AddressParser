namespace Excel.AddressParser

type PositionArray = 
    { topLeft : Position
      bottomRight : Position }
    
    member this.A1Style() = 
        let topLeft = this.topLeft
        let bottomRight = this.bottomRight
        let r1, c1 = topLeft.row, topLeft.column
        let r2, c2 = bottomRight.row, bottomRight.column

        if topLeft = bottomRight then [topLeft.A1Style()]
        elif r1 = 1 && r2 = Row.maxRows then [ (Column.address c1); (Column.address c2)]
        elif c1 = 1 && c2 = Column.maxColumns then [string r1; string r2]
        else [topLeft.A1Style(); bottomRight.A1Style()]
    
    member this.R1C1Style() = 
        let topLeft = this.topLeft
        let bottomRight = this.bottomRight
        let r1, c1 = topLeft.row, topLeft.column
        let r2, c2 = bottomRight.row, bottomRight.column

        if topLeft = bottomRight then [topLeft.R1C1Style()]
        elif r1 = 1 && r2 = Row.maxRows then [sprintf "C%d" c1;sprintf "C%d" c2]
        elif c1 = 1 && c2 = Column.maxColumns then [sprintf "R%d" r1;sprintf "R%d" r2]
        else [topLeft.R1C1Style();bottomRight.R1C1Style()]

    static member EntireRows(r1, r2) = 
        { topLeft = 
              { row = r1
                column = 1 }
          bottomRight = 
              { row = r2
                column = Column.maxColumns } }
    
    static member EntireColumns(c1, c2) = 
        { topLeft = 
              { row = 1
                column = c1 }
          bottomRight = 
              { row = Row.maxRows
                column = c2 } }
    
    static member Entire = 
        { topLeft = 
              { row = 1
                column = 1 }
          bottomRight = 
              { row = Row.maxRows
                column = Column.maxColumns } }
    
    member this.isSubsetOf(that:PositionArray) =
        that.topLeft.isTopLeftOf(this.topLeft) && that.bottomRight.isBottomRightOf(this.bottomRight)

    member this.Move(r,c) =
        let p1 = this.topLeft.move(r,c)
        let p2 = this.bottomRight.move(r,c)
        let par = {topLeft = p1;bottomRight = p2}
        if par.isSubsetOf PositionArray.Entire then
            par
        else
            failwithf "%A" (this,r,c)
    
    static member smartCreate (r1, r2, c1, c2) = 
        let asc a b = 
            if b < a then b, a
            else a, b
        
        let r1, r2 = asc r1 r2
        let c1, c2 = asc c1 c2
        { topLeft = 
              { row = r1
                column = c1 }
          bottomRight = 
              { row = r2
                column = c2 } }
    
    ///当不确定p1，p2的位置时，用此函数
    static member smartCreate (p1, p2) = 
        PositionArray.smartCreate (p1.row, p2.row, p1.column, p2.column)


