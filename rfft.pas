unit rFFT;

{$mode Delphi}

interface
uses
  Classes, SysUtils;


const DataBufLen = 8*1024;
Type PReal64ArrayZeroBased =^TReal64ArrayZeroBased;
     TReal64ArrayZeroBased = array [0..DataBufLen-1] of double;

procedure DoFFT(const re, im: PReal64ArrayZeroBased; const N:integer; const D:integer = -1);
procedure RealFFT(const re: PReal64ArrayZeroBased; var N: integer);


implementation


procedure DoFFT(const re, im: PReal64ArrayZeroBased; const N:integer; const D:integer = -1);
var
  O,I1,M,F,K,I,L,E,J,J1 : integer;
  W,P,Q,R,T,U,V,Z,C,S:double;
begin
  M:=Round(Ln(N)/Ln(2));
  for L:=1 to M do begin
    E:=Round(EXP((M+1-L)*ln(2)));
    F:=E div 2;
    U:=1;
    V:=0;
    Z:=Pi/F;
    c:=cos(z);
    s:=D*sin(z);
    j:=1;
    while j<=f do begin
      I:=J;
      while I<=N do begin
        O:=I+F-1;
        P:=re^[I-1]+re^[o];
        Q:=im^[I-1]+im^[o];
        R:=re^[I-1]-re^[o];
        T:=im^[I-1]-im^[o];
        re^[o]:=R*U-T*V;
        im^[o]:=T*U+R*V;
        re^[I-1]:=P;
        im^[I-1]:=Q;
        inc(I,E);
      end;
      W:=U*C-V*S;
      V:=V*C+U*S;
      U:=W;
      inc(J);
    end;
  end;
  J:=1;
  for I:=1 to N-1 do begin
    if I<J then begin
       J1:=J-1;
       I1:=I-1;
       P:=re^[J1];
       Q:=im^[J1];
       re^[J1]:=re^[I1];
       im^[J1]:=im^[I1];
       re^[I1]:=P;
       im^[I1]:=Q;
    end;
    K:=N div 2;
    while K<J do begin
       J:=J-K;
       K:=K div 2;
    end;
  J:=J+K;
  end;
end;

var im: TReal64ArrayZeroBased;

procedure RealFFT(const re: PReal64ArrayZeroBased; var N: integer);
var count, i: integer;
    scale: double;
begin

  count:=1 shl Round(Ln(N)/Ln(2));


  for i:=0 to count-1 do begin
      im[i]:=0;
  end;
  DoFFT(re, @im, count);
  for i:=0 to count-1 do begin
      im[i]:=0;
  end;
  scale:=2.0/double(count);
  for i:=0 to count div 2 do begin
      re^[i]:= sqrt(sqr(re^[i])+sqr(im[i]))*scale;
  end;
  N:=(count div 2)+1;
end;

end.

