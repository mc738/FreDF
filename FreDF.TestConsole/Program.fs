// Learn more about F# at http://docs.microsoft.com/dotnet/fsharp

open System
open System.Collections
open System.IO
open System.Text.Json
open FreDF.Core
open MigraDocCore.DocumentObjectModel
open MigraDocCore.DocumentObjectModel.MigraDoc.DocumentObjectModel.Shapes
open PdfSharpCore.Utils

// Define a function to construct a message to print
let from whom =
    sprintf "from %s" whom

let testDoc path =
    // Set the image source impl
    ImageSource.ImageSourceImpl <- ImageSharpImageSource<SixLabors.ImageSharp.PixelFormats.Rgba32>() :> ImageSource
    let document = Document()
    let styles = File.ReadAllText "C:\\Users\\44748\\Projects\\PDFBuilder\\styles.json" |> JsonSerializer.Deserialize<PDFStyle>
    styles.DefineDocumentStyle document
    
    
    let pdf = PDF(document)
    pdf.AddHeader1("Header 1", false)
    pdf.AddParagraph("Some text...", false)
    pdf.AddHeader2("Header 2", false)
    pdf.AddParagraph("Some more text", false)
    pdf.AddSection(false)
    pdf.AddImage("C:\\Users\\44748\\Pictures\\Funny-Emo-Meme-Image-Photo-Joke-01.jpeg", "10cm", false, true)
    pdf.Render(path)

let testDoc2 path =
    Pdf.init "C:\\Users\\44748\\Projects\\PDFBuilder\\styles.json"
    |> Pdf.build [
        {
            Portrait = true
            Elements = [
                Elements.h1 "Hello, World!"
                Elements.p "Some text..."
                Elements.h2 "Another header!"
                Elements.p "Some more text..."
            ]
        }
        {
            Portrait = false
            Elements = [
                Elements.h1 "An image!"
                Elements.img "C:\\Users\\44748\\Pictures\\Funny-Emo-Meme-Image-Photo-Joke-01.jpeg" "2cm" false
            ]
        }
    ]
    |> Pdf.render path

[<EntryPoint>]
let main argv =
    // Set the image source impl
    ImageSource.ImageSourceImpl <- ImageSharpImageSource<SixLabors.ImageSharp.PixelFormats.Rgba32>() :> ImageSource
    
    testDoc2 $@"C:\\ProjectData\\TestPDFs\\{DateTime.Now.ToFileTime()}.pdf"
    
    let message = from "F#" // Call the function
    printfn "Hello world %s" message
    0 // return an integer exit code