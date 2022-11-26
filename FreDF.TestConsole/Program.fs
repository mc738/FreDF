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
let from whom = sprintf "from %s" whom

let testDoc path =
    // Set the image source impl
    ImageSource.ImageSourceImpl <- ImageSharpImageSource<SixLabors.ImageSharp.PixelFormats.Rgba32>() :> ImageSource
    let document = Document()

    let styles =
        File.ReadAllText "C:\\Users\\44748\\Projects\\PDFBuilder\\styles.json"
        |> JsonSerializer.Deserialize<PDFStyle>

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
    |> Pdf.build [ { Portrait = true
                     Elements =
                       [ Dsl.h1 "Hello, World!"
                         Dsl.p "Some text..."
                         Dsl.h2 "Another header!"
                         Dsl.p "Some more text..." ] }
                   { Portrait = false
                     Elements =
                       [ Dsl.h1 "An image!"
                         Dsl.img "C:\\Users\\44748\\Pictures\\Funny-Emo-Meme-Image-Photo-Joke-01.jpeg" "2cm" false ] } ]
    |> Pdf.render path

module TemplateTest =

    open FreDF.Core

    let stylePath =
        "C:\\Users\\44748\\Projects\\PDFBuilder\\styles.json"

    let styles =
        File.ReadAllText stylePath
        |> JsonSerializer.Deserialize<PDFStyle>

    let data =
        ({ Data =
            { Values =
                [ "foo",
                  [ [ "bar", Data.Value.Scalar "item 1"
                      "baz",
                      [ [ "inner1", Data.Value.Scalar "Value 1-1"
                          "inner2", Data.Value.Scalar "Value 1-2"
                          "inner3", Data.Value.Scalar "Value 1-3" ]
                        |> Map.ofList
                        |> Data.Value.Object
                        [ "inner1", Data.Value.Scalar "Value 2-1"
                          "inner2", Data.Value.Scalar "Value 2-2"
                          "inner3", Data.Value.Scalar "Value 2-3" ]
                        |> Map.ofList
                        |> Data.Value.Object ]
                      |> Data.Value.Array ]
                    |> Map.ofList
                    |> Data.Value.Object
                    [ "bar", Data.Value.Scalar "item 2"
                      "baz",
                      [ [ "inner1", Data.Value.Scalar "Value 3-1"
                          "inner2", Data.Value.Scalar "Value 3-2"
                          "inner3", Data.Value.Scalar "Value 3-3" ]
                        |> Map.ofList
                        |> Data.Value.Object
                        [ "inner1", Data.Value.Scalar "Value 4-1"
                          "inner2", Data.Value.Scalar "Value 4-2"
                          "inner3", Data.Value.Scalar "Value 4-3" ]
                        |> Map.ofList
                        |> Data.Value.Object ]
                      |> Data.Value.Array ]
                    |> Map.ofList
                    |> Data.Value.Object ]
                  |> Data.Value.Array


                  ]
                |> Map.ofList }
           Iterators = { Items = Map.empty } }: Data.Context)

    let ctx =
        ({ Data = data; Scopes = [] }: Templating.TemplatingContext)

    let template =
        ({ Sections =
            [ ({ Scope =
                  // $.root - name iter1 - Could be defined as $.(iter1)__foo
                  Some(
                      { Name = "iter1"
                        Path =
                          { Root = Data.Root
                            Parts = [ Data.PathPart.Array("foo", "iter1") ] } }: Templating.Scope
                  )
                 ScopeType = Some Templating.SectionScopeType.Inner
                 PageSetup = None
                 Headers = None
                 Footers = None
                 Elements =
                   [ ({ Scope = None
                        Format = None
                        Style = None
                        Elements =
                          [ ({ Content =
                                [ Templating.Content.Value
                                      { Root = Data.Iterator "iter1" // This would be defined as `iter1.bar`
                                        Parts = [ Data.PathPart.Object "bar" ] } ] }: Templating.TextTemplate)
                            |> Templating.ParagraphElementTemplate.Text ] }: Templating.ParagraphTemplate)
                     |> Templating.DocumentElementTemplate.Paragraph
                     ({ Scope = None
                        Borders = None
                        Format = None
                        Shading = None
                        Style = None
                        TopPadding = None
                        BottomPadding = None
                        LeftPadding = None
                        RightPadding = None
                        KeepTogether = None
                        Columns =
                          [ ({ Scope = None
                               Borders = None
                               Format = None
                               Shading = None
                               Style = None
                               Width = Style.Unit.Centimeter 4. |> Some
                               HeadingFormat = None
                               KeepWith = None
                               LeftPadding = None
                               RightPadding = None }: Templating.TableColumnTemplate)
                            ({ Scope = None
                               Borders = None
                               Format = None
                               Shading = None
                               Style = None
                               Width = Style.Unit.Centimeter 4. |> Some
                               HeadingFormat = None
                               KeepWith = None
                               LeftPadding = None
                               RightPadding = None }: Templating.TableColumnTemplate)
                            ({ Scope = None
                               Borders = None
                               Format = None
                               Shading = None
                               Style = None
                               Width = Style.Unit.Centimeter 4. |> Some
                               HeadingFormat = None
                               KeepWith = None
                               LeftPadding = None
                               RightPadding = None }: Templating.TableColumnTemplate) ]
                        Rows =
                          [ ({ Scope = None
                               Borders = None
                               Format = None
                               Shading = None
                               Style = None
                               TopPadding = None
                               BottomPadding = None
                               Height = Style.Unit.Centimeter 1. |> Some
                               HeadingFormat = None
                               KeepWith = None
                               VerticalAlignment = None
                               Cells =
                                 [ ({ Scope = None
                                      Index = 0
                                      Borders = None
                                      Format = None
                                      Shading = None
                                      Style = None
                                      MergeDown = None
                                      MergeRight = None
                                      VerticalAlignment = None
                                      Elements =
                                        [ ({ Scope = None
                                             Format = None
                                             Style = None
                                             Elements =
                                               [ ({ Content = [ Templating.Content.Literal "Field 1" ] }: Templating.TextTemplate)
                                                 |> Templating.ParagraphElementTemplate.Text ] }: Templating.ParagraphTemplate)
                                          |> Templating.CellElementTemplate.Paragraph ] }: Templating.TableCellTemplate)
                                   ({ Scope = None
                                      Index = 1
                                      Borders = None
                                      Format = None
                                      Shading = None
                                      Style = None
                                      MergeDown = None
                                      MergeRight = None
                                      VerticalAlignment = None
                                      Elements =
                                        [ ({ Scope = None
                                             Format = None
                                             Style = None
                                             Elements =
                                               [ ({ Content = [ Templating.Content.Literal "Field 2" ] }: Templating.TextTemplate)
                                                 |> Templating.ParagraphElementTemplate.Text ] }: Templating.ParagraphTemplate)
                                          |> Templating.CellElementTemplate.Paragraph ] }: Templating.TableCellTemplate)
                                   ({ Scope = None
                                      Index = 2
                                      Borders = None
                                      Format = None
                                      Shading = None
                                      Style = None
                                      MergeDown = None
                                      MergeRight = None
                                      VerticalAlignment = None
                                      Elements =
                                        [ ({ Scope = None
                                             Format = None
                                             Style = None
                                             Elements =
                                               [ ({ Content = [ Templating.Content.Literal "Field 3" ] }: Templating.TextTemplate)
                                                 |> Templating.ParagraphElementTemplate.Text ] }: Templating.ParagraphTemplate)
                                          |> Templating.CellElementTemplate.Paragraph ] }: Templating.TableCellTemplate) ] }: Templating.TableRowTemplate)
                            ({ Scope =
                                // iter1.baz - name iter2
                                ({ Name = "iter2"
                                   Path =
                                     { Root = Data.Iterator "iter1"
                                       Parts = [ Data.PathPart.Array("baz", "iter2") ] } }: Templating.Scope)
                                |> Some
                               Borders = None
                               Format = None
                               Shading = None
                               Style = None
                               TopPadding = None
                               BottomPadding = None
                               Height = Style.Unit.Centimeter 1. |> Some
                               HeadingFormat = None
                               KeepWith = None
                               VerticalAlignment = None
                               Cells =
                                 [ ({ Scope = None
                                      Index = 0
                                      Borders = None
                                      Format = None
                                      Shading = None
                                      Style = None
                                      MergeDown = None
                                      MergeRight = None
                                      VerticalAlignment = None
                                      Elements =
                                        [ ({ Scope = None
                                             Format = None
                                             Style = None
                                             Elements =
                                               [ ({ Content =
                                                     [ Templating.Content.Value
                                                           { Root = Data.Iterator "iter2"
                                                             Parts = [ Data.PathPart.Object "inner1" ] } ] }: Templating.TextTemplate)
                                                 |> Templating.ParagraphElementTemplate.Text ] }: Templating.ParagraphTemplate)
                                          |> Templating.CellElementTemplate.Paragraph ] }: Templating.TableCellTemplate)
                                   ({ Scope = None
                                      Index = 1
                                      Borders = None
                                      Format = None
                                      Shading = None
                                      Style = None
                                      MergeDown = None
                                      MergeRight = None
                                      VerticalAlignment = None
                                      Elements =
                                        [ ({ Scope = None
                                             Format = None
                                             Style = None
                                             Elements =
                                               [ ({ Content =
                                                     [ Templating.Content.Value
                                                           { Root = Data.Iterator "iter2"
                                                             Parts = [ Data.PathPart.Object "inner2" ] } ] }: Templating.TextTemplate)
                                                 |> Templating.ParagraphElementTemplate.Text ] }: Templating.ParagraphTemplate)
                                          |> Templating.CellElementTemplate.Paragraph ] }: Templating.TableCellTemplate)
                                   ({ Scope = None
                                      Index = 2
                                      Borders = None
                                      Format = None
                                      Shading = None
                                      Style = None
                                      MergeDown = None
                                      MergeRight = None
                                      VerticalAlignment = None
                                      Elements =
                                        [ ({ Scope = None
                                             Format = None
                                             Style = None
                                             Elements =
                                               [ ({ Content =
                                                     [ Templating.Content.Value
                                                           { Root = Data.Iterator "iter2"
                                                             Parts = [ Data.PathPart.Object "inner3" ] } ] }: Templating.TextTemplate)
                                                 |> Templating.ParagraphElementTemplate.Text ] }: Templating.ParagraphTemplate)
                                          |> Templating.CellElementTemplate.Paragraph ] }: Templating.TableCellTemplate) ] }: Templating.TableRowTemplate) ] }: Templating.TableTemplate)
                     |> Templating.DocumentElementTemplate.Table ] }: Templating.TemplateSection) ] }: Templating.Template)


    let run _ =
        let i = 0

        //let r = Templating.ParagraphTemplate

        let t = template

        let template2 =
            (File.ReadAllText "C:\\ProjectData\\Fiket\\prototypes\\test_template.json"
             |> JsonDocument.Parse)
                .RootElement
            |> Templating.Template.FromJson

        let doc = template2.Build(ctx)

        PdfRenderer.run styles "C:\\ProjectData\\TestPDFs\\template_test-1.pdf" doc


        ()

module FullTemplateTest =

    let stylePath =
        "C:\\Users\\44748\\Projects\\PDFBuilder\\styles.json"

    let styles =
        File.ReadAllText stylePath
        |> JsonSerializer.Deserialize<PDFStyle>

    let run _ =

        let loadJson path =
            (File.ReadAllText path |> JsonDocument.Parse)
                .RootElement

        let unwrap (result: Result<'T, 'E>) =
            match result with
            | Ok r -> r
            | Error _ -> failwith "Error"

        let data =
            Data.Data.FromJson(loadJson "C:\\ProjectData\\Fiket\\prototypes\\test_data.json") |> unwrap

        let ctx =
            ({ Data =
                { Data = data
                  Iterators = { Items = Map.empty } }
               Scopes = [] }: Templating.TemplatingContext)

        let template =
            loadJson "C:\\ProjectData\\Fiket\\prototypes\\test_template.json"
            |> Templating.Template.FromJson

        let doc = template.Build(ctx)

        PdfRenderer.run styles "C:\\ProjectData\\TestPDFs\\template_test-full.pdf" doc
        
module QuoteTemplateTest =

    let stylePath =
        "C:\\Users\\44748\\Projects\\PDFBuilder\\styles.json"

    let styles =
        File.ReadAllText stylePath
        |> JsonSerializer.Deserialize<PDFStyle>

    let run _ =

        let loadJson path =
            (File.ReadAllText path |> JsonDocument.Parse)
                .RootElement

        let unwrap (result: Result<'T, 'E>) =
            match result with
            | Ok r -> r
            | Error _ -> failwith "Error"

        let data =
            Data.Data.FromJson(loadJson "C:\\ProjectData\\Fiket\\prototypes\\test_quote_data.json") |> unwrap

        let ctx =
            ({ Data =
                { Data = data
                  Iterators = { Items = Map.empty } }
               Scopes = [] }: Templating.TemplatingContext)

        let template =
            loadJson "C:\\ProjectData\\Fiket\\prototypes\\test_quote_template.json"
            |> Templating.Template.FromJson

        let doc = template.Build(ctx)

        PdfRenderer.run styles "C:\\ProjectData\\TestPDFs\\quote_test.pdf" doc

[<EntryPoint>]
let main argv =
    let test = Data.Path.Parse("$.foo[iter1]")

    let test2 = Data.Path.Parse("iter1.bar")

    let test3 =
        Templating.Scope.TryParse("$.foo[iter1]")

    let test4 =
        Templating.Scope.TryParse($"iter1.baz[iter2]")

    // Set the image source impl
    ImageSource.ImageSourceImpl <- ImageSharpImageSource<SixLabors.ImageSharp.PixelFormats.Rgba32>() :> ImageSource

    QuoteTemplateTest.run ()

    testDoc2 $@"C:\\ProjectData\\TestPDFs\\{DateTime.Now.ToFileTime()}.pdf"

    let message = from "F#" // Call the function
    printfn "Hello world %s" message
    0 // return an integer exit code
