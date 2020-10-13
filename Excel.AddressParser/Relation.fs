module Excel.AddressParser.Relation

let intersect(a:string)(b:string) =
    let p1 = A1Parser.parse a
    let p2 = A1Parser.parse b
        
    let first = PositionArray.smartCreate(p1.topLeft,p2.topLeft).bottomRight
    let last  = PositionArray.smartCreate(p1.bottomRight,p2.bottomRight).topLeft

    if first.isTopLeftOf last then
        {topLeft = first;bottomRight = last}.A1Style() |> String.concat ":"
    else
        failwithf "'%s'与'%s'不相交" a b

