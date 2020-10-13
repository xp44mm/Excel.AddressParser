///解析A1样式为Position
module Excel.AddressParser.A1Parser

open System.Text.RegularExpressions

///解析A1样式字符串
let (|Cell|Column|Row|) (addr:string) = 
    let patLetters = @"\$?(?<col>[A-Za-z]+)"
    let patDigits = @"\$?(?<row>\d+)"
    let pat = sprintf @"^(%s)?(%s)?$" patLetters patDigits
    let rgx = Regex(pat,RegexOptions.ExplicitCapture)

    let mat = rgx.Match(addr)
    let success (gnm:string) = mat.Groups.[gnm].Success
    let gvalue (gnm:string) = mat.Groups.[gnm].Value
    if mat.Success then
        if success "col" && success "row" then
            let r = int(gvalue "row")
            let c = Column.position(gvalue "col")
            Cell(r,c)
        elif success "col" then
            let c = Column.position(gvalue "col")
            Column(c)
        elif success "row" then
            let r = int(gvalue "row")
            Row(r)
        else
            failwith mat.Value
    else
        failwith mat.Value


///解析单元格A1
let parseToPosition addr = 
    match addr with
    | Cell (r,c) -> {row=r;column=c}
    | _ -> failwithf "%A" addr

///解析范围A1:B2
let parseToPositionArray = 
    function 
    | (Cell _ as c1), (Cell _ as c2) -> {topLeft = parseToPosition c1; bottomRight = parseToPosition c2}
    | Column c1, Column c2 -> PositionArray.EntireColumns(c1, c2)
    | Row r1, Row r2 -> PositionArray.EntireRows(r1, r2)
    | never -> failwithf "%A" never

let parseCells(cells:string list) =
    match cells with
    | [c] -> 
        let p = parseToPosition c
        {topLeft = p; bottomRight = p}
    | [c1;c2] ->
        parseToPositionArray(c1,c2)
    | never -> failwithf "%A" never

let parse(address:string) = address.Split(':') |> Array.toList |> parseCells
