unit Unit1;

interface

uses
  System.Zip, Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Xml.xmldom, Xml.XMLIntf,
  Vcl.ExtCtrls, Vcl.Grids, Vcl.ValEdit, Xml.XMLDoc;

type
  TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    Memo1: TMemo;
    Button1: TButton;
    XMLDocument1: TXMLDocument;
    ValueListEditor1: TValueListEditor;
    Panel1: TPanel;
    XMLDocument2: TXMLDocument;
    Splitter1: TSplitter;
    procedure Button1Click(Sender: TObject);
  private
    FZipFile : TZipFile;
    procedure ReadProperties(const FileName: String);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  ZipStream: TStream;
  LocalHeader: TZipHeader;
  FileName : string;
begin
  if OpenDialog1.Execute then begin
    Memo1.Clear;
    FZipFile := TZipFile.Create;
    try
      FZipFile.Open(OpenDialog1.FileName, TZipMode.zmRead);
//      for fileName in FZipFile.FileNames do
//        Memo1.Lines.Add(FileName);
      FZipFile.Read('_rels/.rels', ZipStream, LocalHeader);
      ZipStream.Position := 0;
      XMLDocument1.LoadFromStream(ZipStream);
      ValueListEditor1.Strings.Clear;
      for var i := 0 to XMLDocument1.DocumentElement.ChildNodes.Count-1 do begin
        var XmlNode := XMLDocument1.DocumentElement.ChildNodes.Nodes[i];
        var AttType := ExtractFileName(StringReplace(XmlNode.Attributes['Type'],
          '/','\', [rfReplaceAll]));
        if AnsiSameText(AttType, 'core-properties') or
             AnsiSameText(AttType, 'extended-properties') then
          ReadProperties(XmlNode.Attributes['Target']);
      end;
      Memo1.Lines.Text := FormatXmlData(XmlDocument1.XML.Text);
    finally
      FZipFile.Free;
    end;
  end;
end;

procedure TForm1.ReadProperties(const FileName : String);
var
  ZipStream : TStream;
  LocalHeader: TZipHeader;
  i : Integer;
  XmlNode : IXMLNode;
begin
  FZipFile.Read(FileName, ZipStream, LocalHeader);
  ZipStream.Position := 0;
  XMLDocument2.LoadFromStream(ZipStream);
  // Read the properties
  for i := 0 to XMLDocument2.DocumentElement.ChildNodes.Count-1 do begin
    XmlNode := XMLDocument2.DocumentElement.ChildNodes.Nodes[i];
    try
      // Found new property
      ValueListEditor1.InsertRow(XmlNode.NodeName,XmlNode.NodeValue, True);
    except
      // Property is not simple - discard it.
      On EXMLDocError do
        ;
      // Property is null - add empty string
      On EVariantTypeCastError do
        ValueListEditor1.InsertRow(XmlNode.NodeName,'', True);
    end;
  end;
end;

end.
