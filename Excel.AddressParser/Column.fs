[<RequireQualifiedAccess>]
module Excel.AddressParser.Column

open System

///大写字母表
let private uppercases = [| 'A'..'Z' |]

///小写字母表
let private lowercases = [| 'a'..'z' |]

///字母，位置，查詢表
let private lookup =
    [
        yield! uppercases |> Array.mapi(fun i c ->c,i+1)
        yield! lowercases |> Array.mapi(fun i c ->c,i+1)
    ]|> Map.ofList

///字符在26进制中代表的值，字符不区分大小写
let valueof ch = lookup.[ch]

///工作表最大列地址
[<Literal>]
let maxLetters = "XFD" //letter maxColumn

/// = maxLetters.Length
[<Literal>]
let maxLength = 3

/// = maxLetters.ToCharArray() |> Array.map(valueof)
let maxPositions = [|24; 6; 4|]

///判断字母序列是否为合法列地址，不区分大小写
let isColumn (letters: string) =
    match compare letters.Length maxLength with
    | -1 -> true
    |  0 ->
        //应用valueof，以忽略大小写
        let positions = letters.ToCharArray() |> Array.map(valueof)
        positions <= maxPositions
    | 1 -> false
    | never -> failwith ""

///求字母序列的列位置，字母不区分大小写，位置基于1而非0。
let position(letters : string) =
    letters.ToCharArray()
    |> Array.map(valueof)
    |> Array.rev
    |> Array.mapi(fun i n -> (pown 26 i) * n)
    |> Array.sum

/// = position maxLetters
[<Literal>]
let maxColumns = 16384

///由位置得到地址：Column -> Address（大写字母）
let address num =
    let rec loop value =
        [|
            yield value % 26
            let high = value / 26
            if high > 0 then
                yield! loop high
        |]

    loop num
    |> Array.rev
    |> Array.map(fun i -> uppercases.[i-1])
    |> fun chars -> String chars

///positionOfA n : n个A对应的Excel位置，Excel位置基于1而非基于0。
///positionOfA 1 = pos of A   =   1;
///positionOfA 2 = pos of AA  =  27;
///positionOfA 3 = pos of AAA = 703;
//let positionOfA n =
//    String.replicate n "A"
//    |> position

//let address num =
//    //确定生成字母的长度
//    let offset, length =
//        let rec loop len =
//            match num - positionOfA len with
//            | offset when offset > -1 -> offset, len
//            | _ -> loop (len - 1)
//        loop maxLength

//    let rec loop value =
//        seq {
//            yield uppercases.[value % 26]
//            yield! loop (value / 26)
//        }

//    let chars =
//        loop offset
//        |> Seq.take length
//        |> Array.ofSeq
//        |> Array.rev

    //String chars

//需要測試是否和position等價
//let position(letters : string) =
//    let len = letters.Length
//    let offset =
//        letters.ToCharArray()
//        |> Array.map(valueof)
//        |> Array.mapi(fun i value ->
//            value * pown 26 (len - i - 1) //求每个字符的数值
//        )
//        |> Array.sum
//    positionOfA letters.Length + offset
//let private positionOfA =
//    let amounts =
//        let rec loop acc last =
//            seq {
//                let acc = acc + last
//                yield acc
//                yield! loop acc (26 * last)
//            }
//        loop 0 1
//        |> Seq.take maxLength
//        |> Seq.toArray
//    fun bit -> amounts.[bit - 1]
