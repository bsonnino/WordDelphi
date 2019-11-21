unit Unit3;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  Xml.xmldom, Xml.XMLIntf, Xml.XmlDoc, System.Zip;

type
  TForm3 = class(TForm)
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    FZipFile: TZipFile;
    function CreateContentTypes: TStream;
    function CreateRels: TStream;
    procedure AddFormattedText(Body: IXMLNode; FontName: String);
    function CreateDocument: TStream;
    procedure SetPageFormat(Body: IXMLNode);
    function CreateHeader: TStream;
    procedure CreateXmlHeader(XmlDocument: TXmlDocument);
    function SaveXmlDocToStream(XmlDocument: TXmlDocument): TStream;
    function CreateDocRels: TStream;
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

procedure TForm3.Button1Click(Sender: TObject);
begin
  FZipFile := TZipFile.Create;
  try
    FZipFile.Open('ComplexFile.docx', TZipMode.zmWrite);
    var ContentTypeStream := CreateContentTypes();
    try
      FZipFile.Add(ContentTypeStream, '[Content_Types].xml');
    finally
      ContentTypeStream.Free;
    end;
    var RelsStream := CreateRels();
    try
      FZipFile.Add(RelsStream, '_rels\.rels');
    finally
      RelsStream.Free;
    end;
    RelsStream := CreateDocRels();
    try
      FZipFile.Add(RelsStream, 'word\_rels\document.xml.rels');
    finally
      RelsStream.Free;
    end;
    var DocumentStream := CreateDocument();
    try
      FZipFile.Add(DocumentStream, 'word\document.xml');
    finally
      DocumentStream.Free;
    end;
    var HeaderStream := CreateHeader();
    try
      FZipFile.Add(HeaderStream, 'word\header1.xml');
    finally
      DocumentStream.Free;
    end;
    FZipFile.Close;
  finally
    FZipFile.Free;
  end;
end;

function TForm3.SaveXmlDocToStream(XmlDocument : TXmlDocument) : TStream;
begin
  Result := TMemoryStream.Create;
  XmlDocument.SaveToStream(Result);
  Result.Position := 0;
end;

function TForm3.CreateContentTypes: TStream;
begin
  var XmlDocument := TXmlDocument.Create(self);
  try
    CreateXmlHeader(XmlDocument);
    // Root Node
    var Root := XmlDocument.addChild('Types',
      'http://schemas.openxmlformats.org/package/2006/content-types');
    // Type Definition
    var TypeNode := Root.addChild('Default');
    TypeNode.Attributes['Extension'] := 'rels';
    TypeNode.Attributes['ContentType'] :=
      'application/vnd.openxmlformats-package.relationships+xml';
    TypeNode := Root.addChild('Default');
    TypeNode.Attributes['Extension'] := 'xml';
    TypeNode.Attributes['ContentType'] :=
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
    TypeNode := Root.AddChild('Override');
    TypeNode.Attributes['PartName'] := '/word/header1.xml';
    TypeNode.Attributes['ContentType'] := 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
    Result := SaveXmlDocToStream(XmlDocument);
  finally
    XmlDocument.Free
  end;
end;

procedure TForm3.CreateXmlHeader(XmlDocument : TXmlDocument);
begin
  XmlDocument.Options := [doNodeAutoIndent];
  XmlDocument.Active := True;
  XmlDocument.Encoding := 'UTF-8';
  XmlDocument.Version := '1.0';
  XmlDocument.StandAlone := 'yes';
end;

function TForm3.CreateRels: TStream;
begin
  var XmlDocument := TXmlDocument.Create(self);
  try
    CreateXmlHeader(XmlDocument);
    // Root Node
    var Root := XmlDocument.addChild('Relationships',
      'http://schemas.openxmlformats.org/package/2006/relationships');
    // Relationship Definition
    var Rel := Root.addChild('Relationship');
    Rel.Attributes['Id'] := 'rId1';
    Rel.Attributes['Type'] :=
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
    Rel.Attributes['Target'] := 'word/document.xml';
    Result := SaveXmlDocToStream(XmlDocument);
  finally
    XmlDocument.Free;
  end;
end;

function TForm3.CreateDocRels : TStream;
begin
  var XmlDocument := TXmlDocument.Create(self);
  try
    CreateXmlHeader(XmlDocument);
    // Root node
    var Root := XMLDocument.addChild('Relationships',
      'http://schemas.openxmlformats.org/package/2006/relationships');
    // Relationship definitions
    var Rel := Root.AddChild('Relationship');
    Rel.Attributes['Id'] := 'rId1';
    Rel.Attributes['Type'] :=
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
    Rel.Attributes['Target'] := 'header1.xml';
    Result := SaveXmlDocToStream(XmlDocument);
  finally
    XmlDocument.Free;
  end;
end;

function TForm3.CreateDocument: TStream;
begin
  var XmlDocument := TXmlDocument.Create(self);
  try
    CreateXmlHeader(XmlDocument);
    // Root node
    var Root := XMLDocument.addChild('w:wordDocument');
    Root.DeclareNamespace('w',
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
    Root.DeclareNamespace('r',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
    var Body := Root.addChild('w:body');
    for var i := 0 to Screen.Fonts.Count - 1 do begin
      AddFormattedText(Body, Screen.Fonts[i]);
    end;
    SetPageFormat(Body);
    Result := SaveXmlDocToStream(XmlDocument);
  finally
    XmlDocument.Free;
  end;
end;

procedure TForm3.AddFormattedText(Body: IXMLNode; FontName: String);
begin
  var Run := Body.addChild('w:p').addChild('w:r');
  var RunPr := Run.addChild('w:rPr');
  var Font := RunPr.addChild('w:rFonts');
  Font.Attributes['w:ascii'] := FontName;
  Font.Attributes['w:hAnsi'] := FontName;
  Font.Attributes['w:cs'] := FontName;
  RunPr.addChild('w:sz').Attributes['w:val'] := 30;
  Run.addChild('w:t').NodeValue := FontName;
  Run.addChild('w:tab');
  Run.addChild('w:t').NodeValue :=
    'The quick brown fox jumps over the lazy dog';
end;

procedure TForm3.SetPageFormat(Body: IXMLNode);
begin
  var SectPr := Body.AddChild('sectPr');
  var Header := SectPr.AddChild('w:headerReference');
  Header.Attributes['w:type'] := 'default';
  Header.Attributes['r:id'] := 'rId1';
  var PgSz := SectPr.AddChild('w:pgSz');
  PgSz.Attributes['w:w'] := Round(297/25.4*1440);
  PgSz.Attributes['w:h'] := Round(210/25.4*1440);
  PgSz := SectPr.AddChild('w:pgMar');
  PgSz.Attributes['w:top'] := 1440;
  PgSz.Attributes['w:bottom'] := 1440;
  PgSz.Attributes['w:left'] := 720;
  PgSz.Attributes['w:right'] := 720;
  PgSz.Attributes['w:header'] := 720;
  PgSz.Attributes['w:footer'] := 720;
end;

function TForm3.CreateHeader : TStream;
begin
  var XmlDocument := TXmlDocument.Create(self);
  try
    CreateXmlHeader(XmlDocument);
    // Root node
    var Root := XmlDocument.addChild('w:hdr');
    Root.DeclareNamespace('w',
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
    var Header := Root.addChild('w:p');
    Header.addChild('w:r').addChild('w:t').NodeValue := 'Text 1';
    var PTab := Header.addChild('w:r').addChild('w:ptab');
    PTab.Attributes['w:relativeTo'] := 'margin';
    PTab.Attributes['w:alignment'] := 'center';
    PTab.Attributes['w:leader'] := 'none';
    Header.addChild('w:r').addChild('w:t').NodeValue := 'Text 2';
    PTab := Header.addChild('w:r').addChild('w:ptab');
    PTab.Attributes['w:relativeTo'] := 'margin';
    PTab.Attributes['w:alignment'] := 'right';
    PTab.Attributes['w:leader'] := 'none';
    Header.addChild('w:fldSimple').Attributes['w:instr'] :=
      'PAGE \* MERGEFORMAT';
    Result := SaveXmlDocToStream(XmlDocument);
  finally
    XmlDocument.Free;
  end;
end;

end.
