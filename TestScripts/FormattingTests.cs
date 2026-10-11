// Regression tests for RichEditBox.ToggleSuperscript / ToggleSubscript / ToggleList.
// A .NET 10 file-based app that is not part of the solution. Run it from the repository root with:
//   dotnet run TestScripts/FormattingTests.cs
// The editor is never shown. Prints one line per check; the exit code is the number of failed checks.
#:property TargetFramework=net10.0-windows
#:property UseWindowsForms=true
#:property PublishAot=false
#:project ../RichEditorNET/RichEditorNET.csproj

using EllipticBit.RichEditorNET;
using EllipticBit.RichEditorNET.Formatting;
using EllipticBit.RichEditorNET.TextObjectModel2;

int failed = 0;
// RichEdit's OLE interfaces need an STA thread; top-level statements run on an MTA thread.
var thread = new Thread(() => failed = RunTests());
thread.SetApartmentState(ApartmentState.STA);
thread.Start();
thread.Join();
return failed;

static int RunTests() {
	using var ed = new RichEditBox();
	_ = ed.Handle; // creates the (hidden) window, which sets up TextDocument
	var doc = ed.TextDocument;
	int pass = 0, fail = 0;

	void Check(string name, bool ok, string detail) {
		if (ok) pass++; else fail++;
		Console.WriteLine($"{(ok ? "PASS" : "FAIL")}  {name}  [{detail}]");
	}

	(int Sup, int Sub) State(int cp) {
		var font = doc.Range2(cp, cp + 1).Font;
		return (font.Superscript, font.Subscript);
	}

	const int on = tomConstants.tomTrue, off = tomConstants.tomFalse;

	// Superscript / subscript
	ed.Text = "mc2 H2O";
	ed.Select(2, 1);
	ed.ToggleSuperscript();
	Check("ToggleSuperscript turns superscript on", State(2) == (on, off), State(2).ToString());
	ed.ToggleSuperscript();
	Check("ToggleSuperscript again turns it off", State(2) == (off, off), State(2).ToString());
	ed.Select(5, 1);
	ed.ToggleSubscript();
	Check("ToggleSubscript turns subscript on", State(5) == (off, on), State(5).ToString());
	ed.ToggleSubscript();
	Check("ToggleSubscript again turns it off", State(5) == (off, off), State(5).ToString());
	ed.ToggleSubscript();
	ed.ToggleSuperscript();
	Check("ToggleSuperscript on subscript text switches to superscript", State(5) == (on, off), State(5).ToString());
	ed.ToggleSubscript();
	Check("ToggleSubscript on superscript text switches to subscript", State(5) == (off, on), State(5).ToString());
	ed.Select(2, 1);
	ed.ToggleSuperscript();
	Check("Html contains <sup> and <sub>", ed.Html.Contains("mc<sup>2</sup>") && ed.Html.Contains("H<sub>2</sub>O"), ed.Html);

	// Lists: the numbering RichEdit displays, read with tomIncludeNumbering
	foreach (var style in Enum.GetValues<ListStyle>()) {
		ed.Text = "one\ntwo\nthree";
		ed.SelectAll();
		ed.ToggleList(style);
		if (style == ListStyle.Bullet) {
			Check("ToggleList(Bullet) creates a bullet list", ed.Html.StartsWith("<ul><li>one</li>"), ed.Html);
		}
		else {
			string numbered = doc.Range2(0, ed.TextLength).GetText2(tomConstants.tomIncludeNumbering)
				.Replace("\r", " | ").Replace("\t", " ").Trim();
			string first = style switch {
				ListStyle.Decimal => "1.",
				ListStyle.LowerAlpha => "a.",
				ListStyle.UpperAlpha => "A.",
				ListStyle.LowerRoman => "i.",
				_ => "I."
			};
			Check($"ToggleList({style}) numbers from {first}", numbered.StartsWith(first), numbered);
		}
		ed.SelectAll();
		ed.ToggleList(style);
		Check($"ToggleList({style}) again removes the list",
			(doc.Range2(0, 0).Para.ListType & 0xFFFF) == tomConstants.tomListNone && ed.Html.StartsWith("<p>one</p>"), ed.Html);
	}

	Console.WriteLine($"{pass} passed, {fail} failed");
	return fail;
}
