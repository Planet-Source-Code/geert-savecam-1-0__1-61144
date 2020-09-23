library WickPic;
uses
  Windows,
  SysUtils,
  Classes,
  Dialogs,
  Graphics,
  Jpeg;

Function C2P(B_bestand : PChar;
                 J_bestand : PChar;
                        K : Integer) : Boolean; StdCall;
var
  Bitmap : TBitmap;
  Jpg : TJpegImage;
begin
  Bitmap := TBitmap.Create;
  Jpg := TJpegImage.Create;
  try
    if (K < 1) or (K > 100) then K := 100;
    Jpg.CompressionQuality := K;
    Jpg.Compress;
    Bitmap.LoadFromFile(B_bestand);
    Jpg.Assign(Bitmap);
    Jpg.SaveToFile(ChangeFileExt(J_bestand, '.jpg'));
    Jpg.Free;
    Bitmap.Free;
    Result := True;
  except
    Result := False;
   end;
end;



EXPORTS
  C2P;

begin

end.
