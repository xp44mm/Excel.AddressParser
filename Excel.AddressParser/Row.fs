[<RequireQualifiedAccess>]
module Excel.AddressParser.Row

//工作表允许的最大行数
[<Literal>]
let maxRows = 1048576

///最大行号的数字位数
let maxrowDigits = maxRows.ToString().Length

let isRow (number:string) =
    if 0 < number.Length && number.Length < maxrowDigits then
        true
    elif number.Length = maxrowDigits && int number < maxRows then
        true
    else
        false
        