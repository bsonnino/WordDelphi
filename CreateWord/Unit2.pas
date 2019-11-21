unit Unit2;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Xml.xmldom, Xml.XMLIntf, Xml.XmlDoc,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, System.Zip;

type
  TForm2 = class(TForm)
    Memo1: TMemo;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
    FZipFile: TZipFile;
    function CreateContentTypes: TStream;
    function CreateRels: TStream;
    function CreateDocument: TStream;
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

procedure TForm2.Button1Click(Sender: TObject);
begin
  FZipFile := TZipFile.Create;
  try
    FZipFile.Open('SimpleFile.docx', TZipMode.zmWrite);
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
    var DocumentStream := CreateDocument();
    try
      FZipFile.Add(DocumentStream, 'word\document.xml');
    finally
      DocumentStream.Free;
    end;
    FZipFile.Close;
  finally
    FZipFile.Free;
  end;
end;

function TForm2.CreateContentTypes: TStream;
begin
  var XmlDocument := TXmlDocument.Create(self);
  try
    XmlDocument.Options := [doNodeAutoIndent];
    XmlDocument.Active := True;
    // Heading
    XmlDocument.Encoding := 'UTF-8';
    XmlDocument.Version := '1.0';
    XmlDocument.StandAlone := 'yes';
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
    // Save on output stream
    Result := TMemoryStream.Create;
    XMLDocument.SaveToStream(Result);
    Result.Position := 0;
  finally
    XmlDocument.Free
  end;
end;

function TForm2.CreateRels: TStream;
var
  Root: IXmlNode;
  Rel: IXmlNode;
begin
  var XmlDocument := TXmlDocument.Create(self);
  try
    XmlDocument.Options := [doNodeAutoIndent];
    XmlDocument.Active := True;
    // Heading
    XmlDocument.Encoding := 'UTF-8';
    XmlDocument.Version := '1.0';
    XmlDocument.StandAlone := 'yes';
    // Root Node
    Root := XMLDocument.addChild('Relationships',
      'http://schemas.openxmlformats.org/package/2006/relationships');
    // Relationship Definition
    Rel := Root.addChild('Relationship');
    Rel.Attributes['Id'] := 'rId1';
    Rel.Attributes['Type'] :=
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
    Rel.Attributes['Target'] := 'word/document.xml';
    // Save on output stream
    Result := TMemoryStream.Create;
    XmlDocument.SaveToStream(Result);
    Result.Position := 0;
  finally
    XmlDocument.Free;
  end;
end;

function TForm2.CreateDocument: TStream;
var
  Root: IXmlNode;
begin
  var XmlDocument := TXmlDocument.Create(self);
  try
    XmlDocument.Options := [doNodeAutoIndent];
    XmlDocument.Active := True;
    // Heading
    XmlDocument.Encoding := 'UTF-8';
    XmlDocument.Version := '1.0';
    XmlDocument.StandAlone := 'yes';
    // Root Node
    Root := XmlDocument.addChild('wordDocument',
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
    // Save Text
    Root.addChild('body').addChild('p').addChild('r').addChild('t').NodeValue :=
      Memo1.Text;
    // Save on Output Stream
    Result := TMemoryStream.Create;
    XmlDocument.SaveToStream(Result);
    Result.Position := 0;
  finally
    XmlDocument.Free;
  end;
end;

end.
