using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using EllipticBit.RichEditorNET.Formatting;
using EllipticBit.RichEditorNET.Support;
using EllipticBit.RichEditorNET.TextObjectModel2;

namespace EllipticBit.RichEditorNET
{
	public class RichEditBox : RichTextBox
	{
		static RichEditBox() {
			PInvoke.LoadLibrary(PInvoke.MSFTEDIT_DLL);
		}

		public RichEditBox() {
			HideSelection = false;
			AcceptsTab = true;
			Multiline = true;
			ScrollBars = RichTextBoxScrollBars.Vertical;
		}

		protected override CreateParams CreateParams {
			get {
				var cp = base.CreateParams;
				cp.ClassName = PInvoke.RICHEDIT_CLASS;
				return cp;
			}
		}

		#region - Properties -

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables the system spell checker for this editor.")]
		public bool EnableSpellCheck {
			get => _enableSpellCheck;
			set {
				_enableSpellCheck = value;
				if (IsHandleCreated) {
					ApplySpellCheckOption();
				}
			}
		}

		private bool _enableSpellCheck = true;

		[DefaultValue(320)]
		[Category("Behavior")]
		[Description("Default display width in pixels for inserted images.")]
		public int DefaultImageWidth {
			get => _defaultImageWidth;
			set {
				if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value), "Width must be greater than zero.");
				_defaultImageWidth = value;
			}
		}

		private int _defaultImageWidth = 320;

		[DefaultValue(240)]
		[Category("Behavior")]
		[Description("Default display height in pixels for inserted images.")]
		public int DefaultImageHeight {
			get => _defaultImageHeight;
			set {
				if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value), "Height must be greater than zero.");
				_defaultImageHeight = value;
			}
		}

		private int _defaultImageHeight = 240;

		[DefaultValue(false)]
		[Category("Behavior")]
		[Description("Enables CommonMark Markdown mode. Disables formatting options not available in CommonMark.")]
		public bool EnableCommonMarkdown {
			get => _enableCommonMarkdown;
			set {
				_enableCommonMarkdown = value;
				if (!value) _enableGithubMarkdown = false;
			}
		}

		private bool _enableCommonMarkdown;

		[DefaultValue(false)]
		[Category("Behavior")]
		[Description("Enables GitHub Flavored Markdown mode. Implies EnableCommonMarkdown. Disables formatting options not available in GitHub Flavored Markdown.")]
		public bool EnableGithubMarkdown {
			get => _enableGithubMarkdown;
			set {
				_enableGithubMarkdown = value;
				if (value) _enableCommonMarkdown = true;
			}
		}

		private bool _enableGithubMarkdown;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables bold formatting.")]
		public bool EnableBold { get; set; } = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables italic formatting.")]
		public bool EnableItalic { get; set; } = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables underline formatting.")]
		public bool EnableUnderline {
			get => _enableUnderline && !_enableCommonMarkdown;
			set => _enableUnderline = value;
		}

		private bool _enableUnderline = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables strikethrough formatting.")]
		public bool EnableStrikeThrough {
			get => _enableStrikeThrough && (!_enableCommonMarkdown || _enableGithubMarkdown);
			set => _enableStrikeThrough = value;
		}

		private bool _enableStrikeThrough = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables font color formatting.")]
		public bool EnableFontColor {
			get => _enableFontColor && !_enableCommonMarkdown;
			set => _enableFontColor = value;
		}

		private bool _enableFontColor = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables background color formatting.")]
		public bool EnableBackgroundColor {
			get => _enableBackgroundColor && !_enableCommonMarkdown;
			set => _enableBackgroundColor = value;
		}

		private bool _enableBackgroundColor = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables font name formatting.")]
		public bool EnableFontName {
			get => _enableFontName;
			set => _enableFontName = value;
		}

		private bool _enableFontName = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables font size formatting.")]
		public bool EnableFontSize {
			get => _enableFontSize && !EnableHtmlFontSizing;
			set => _enableFontSize = value;
		}

		private bool _enableFontSize = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables superscript formatting.")]
		public bool EnableSuperscript {
			get => _enableSuperscript && !_enableCommonMarkdown;
			set => _enableSuperscript = value;
		}

		private bool _enableSuperscript = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables subscript formatting.")]
		public bool EnableSubscript {
			get => _enableSubscript && !_enableCommonMarkdown;
			set => _enableSubscript = value;
		}

		private bool _enableSubscript = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables list formatting.")]
		public bool EnableLists { get; set; } = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables paragraph alignment formatting.")]
		public bool EnableAlignment {
			get => _enableAlignment && !_enableCommonMarkdown;
			set => _enableAlignment = value;
		}

		private bool _enableAlignment = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables paragraph indentation.")]
		public bool EnableIndent {
			get => _enableIndent && !_enableCommonMarkdown;
			set => _enableIndent = value;
		}

		private bool _enableIndent = true;

		[DefaultValue(false)]
		[Category("Behavior")]
		[Description("Enables HTML heading sizes instead of free-form font size/name. Automatically enabled when markdown modes are active.")]
		public bool EnableHtmlFontSizing {
			get => _enableHtmlFontSizing || _enableCommonMarkdown;
			set => _enableHtmlFontSizing = value;
		}

		private bool _enableHtmlFontSizing;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables hyperlink insertion.")]
		public bool EnableHyperlinks { get; set; } = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables image insertion.")]
		public bool EnableImages { get; set; } = true;

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("When true, the popup toolbar shows the embedded image button. When false, it shows the linked thumbnail button.")]
		public bool RequireEmbeddedImage { get; set; } = true;

		[DefaultValue(false)]
		[Category("Behavior")]
		[Description("When true, loading HTML throws NotSupportedException for any unsupported HTML tag, attribute, or CSS property. When false, unsupported HTML is silently ignored.")]
		public bool EnableStrictHtml { get; set; }

		[DefaultValue(false)]
		[Category("Behavior")]
		[Description("When true, the Html property getter returns a complete HTML document with html and body tags. When false, returns only the document fragment.")]
		public bool ExportCompleteHtmlDocument { get; set; }

		[DefaultValue(HtmlImageFormat.Png)]
		[Category("Behavior")]
		[Description("Specifies the image format to use when embedding images in HTML output. Defaults to PNG.")]
		public HtmlImageFormat HtmlThumbnailImageFormat { get; set; } = HtmlImageFormat.Png;

		[DefaultValue(false)]
		[Category("Behavior")]
		[Description("When true, prevents the popup editing toolbar from being displayed.")]
		public bool DisableEditingPopup { get; set; }

		[DefaultValue(null)]
		[Category("Behavior")]
		[Description("Contains string content from remote sources. When set and different from the currently displayed text, the control background turns orange to indicate external changes.")]
		public string? RemoteContent {
			get => _remoteContent;
			set {
				_remoteContent = value;
				UpdateRemoteContentHighlight();
			}
		}

		private string? _remoteContent;
		private Color _savedBackColor;
		private bool _remoteContentHighlighted;

		/// <summary>
		/// Occurs when the Insert Hyperlink button is clicked on the popup toolbar.
		/// </summary>
		[Category("Action")]
		[Description("Occurs when the Insert Hyperlink button is clicked on the popup toolbar.")]
		public event EventHandler? InsertHyperlinkClicked;

		/// <summary>
		/// Occurs when the Insert Image button is clicked on the popup toolbar.
		/// </summary>
		[Category("Action")]
		[Description("Occurs when the Insert Image button is clicked on the popup toolbar.")]
		public event EventHandler? InsertImageClicked;

		/// <summary>
		/// Occurs when the Insert Linked Thumbnail button is clicked on the popup toolbar.
		/// </summary>
		[Category("Action")]
		[Description("Occurs when the Insert Linked Thumbnail button is clicked on the popup toolbar.")]
		public event EventHandler? InsertLinkedThumbnailClicked;

		private ITextDocument2? _textDocument;
		private IRichEditOle? _richEditOle;
		private RichEditOleCallback? _oleCallback;
		private PopupToolbar? _activeToolbar;
		private int _activeToolbarDpi;
		private static readonly Dictionary<int, ToolbarIconCache> _iconCaches = new();
		private bool _spellCheckMenuShown;
		private string? _pendingHtml;
		private string? _pendingMarkdown;

		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public ITextDocument2 TextDocument => _textDocument ?? throw new InvalidOperationException("The control handle has not been created.");

		/// <summary>
		/// Gets or sets the document content as a markdown string. The flavor is determined by
		/// <see cref="EnableCommonMarkdown"/> and <see cref="EnableGithubMarkdown"/>.
		/// </summary>
		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public string Markdown {
			get {
				if (_textDocument == null) return _pendingMarkdown ?? string.Empty;
				return MarkdownFormatter.ToMarkdown(TextDocument, EnableGithubMarkdown);
			}
			set {
				this.Clear();
				//If value is empty then just clear the box and reset undo so the empty state is the baseline.
				if (string.IsNullOrEmpty(value)) {
					this.ClearUndo();
					return;
				}

				if (_textDocument == null) {
					_pendingMarkdown = value;
					_pendingHtml = null;
					return;
				}

				bool priorROValue = this.ReadOnly;
				this.ReadOnly = false;

				MarkdownFormatter.FromMarkdown(TextDocument, value, EnableGithubMarkdown);

				this.ReadOnly = priorROValue;

				// Reset the undo queue so the populated content is the first/baseline state
				// and Ctrl-Z cannot wipe it out by undoing the load operations themselves.
				this.ClearUndo();
			}
		}

		/// <summary>
		/// Gets or sets the document content as a html string.
		/// </summary>
		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public string Html {
			get {
				if (_textDocument == null) return _pendingHtml ?? string.Empty;
				return HtmlFormatter.ToHtml(TextDocument, Handle, EnableHtmlFontSizing, ExportCompleteHtmlDocument, Font.Name, Font.SizeInPoints, HtmlThumbnailImageFormat);
			}
			set {
				this.Clear();
				//If value is empty then just clear the box and reset undo so the empty state is the baseline.
				if (string.IsNullOrEmpty(value)) {
					this.ClearUndo();
					return;
				}

				if (_textDocument == null) {
					_pendingHtml = value;
					_pendingMarkdown = null;
					return;
				}

				bool priorROValue = this.ReadOnly;
				this.ReadOnly = false;

				HtmlFormatter.FromHtml(TextDocument, _richEditOle, value, EnableStrictHtml);

				this.ReadOnly = priorROValue;

				// Reset the undo queue so the populated content is the first/baseline state
				// and Ctrl-Z cannot wipe it out by undoing the load operations themselves.
				this.ClearUndo();
			}
		}

		internal void RaiseInsertHyperlinkClicked() => InsertHyperlinkClicked?.Invoke(this, EventArgs.Empty);
		internal void RaiseInsertImageClicked() => InsertImageClicked?.Invoke(this, EventArgs.Empty);
		internal void RaiseInsertLinkedThumbnailClicked() => InsertLinkedThumbnailClicked?.Invoke(this, EventArgs.Empty);

		/// <summary>
		/// Returns a TOM2 range corresponding to the current WinForms selection.
		/// Unlike <see cref="ITextDocument2.Selection2"/>, this works even when the control does not have focus.
		/// The caller must release the returned object via <see cref="Marshal.ReleaseComObject"/>.
		/// </summary>
		private ITextRange2 GetSelectionRange() => TextDocument.Range2(SelectionStart, SelectionStart + SelectionLength);

		#endregion

		#region - Control Behaviors -

		protected override void OnHandleCreated(EventArgs e) {
			base.OnHandleCreated(e);

			IntPtr pOle = IntPtr.Zero;
			PInvoke.SendMessage(Handle, PInvoke.EM_GETOLEINTERFACE, IntPtr.Zero, ref pOle);
			if (pOle != IntPtr.Zero) {
				_textDocument = Marshal.GetObjectForIUnknown(pOle) as ITextDocument2;
				_richEditOle = Marshal.GetObjectForIUnknown(pOle) as IRichEditOle;
				Marshal.Release(pOle);
			}

			_oleCallback = new RichEditOleCallback();
			IntPtr pCallback = Marshal.GetIUnknownForObject(_oleCallback);
			try {
				PInvoke.SendMessage(Handle, PInvoke.EM_SETOLECALLBACK, IntPtr.Zero, pCallback);
			}
			finally {
				Marshal.Release(pCallback);
			}

			PInvoke.SendMessage(Handle, PInvoke.EM_SETOPTIONS, (IntPtr)PInvoke.ECOOP_AND, (IntPtr)(~PInvoke.ECO_AUTOWORDSELECTION));
			ApplySpellCheckOption();

			bool priorROValue = this.ReadOnly;
			this.ReadOnly = false;

			bool loadedPending = false;
			if (_pendingHtml != null) {
				HtmlFormatter.FromHtml(TextDocument, _richEditOle, _pendingHtml, EnableStrictHtml);
				_pendingHtml = null;
				loadedPending = true;
			}
			else if (_pendingMarkdown != null) {
				MarkdownFormatter.FromMarkdown(TextDocument, _pendingMarkdown, EnableGithubMarkdown);
				_pendingMarkdown = null;
				loadedPending = true;
			}

			this.ReadOnly = priorROValue;

			// Ensure pending content set before the handle existed becomes the undo baseline.
			if (loadedPending) this.ClearUndo();
		}

		protected override void OnHandleDestroyed(EventArgs e) {
			if (_activeToolbar != null) {
				_activeToolbar.Dispose();
				_activeToolbar = null;
			}

			if (_richEditOle != null) {
				Marshal.ReleaseComObject(_richEditOle);
				_richEditOle = null;
			}

			if (_textDocument != null) {
				Marshal.ReleaseComObject(_textDocument);
				_textDocument = null;
			}

			_oleCallback = null;

			base.OnHandleDestroyed(e);
		}

		protected override void OnDoubleClick(EventArgs e) {
			base.OnDoubleClick(e);

			string selected = SelectedText;
			if (selected.Length > 0 && char.IsWhiteSpace(selected[selected.Length - 1])) {
				base.SelectionLength = selected.TrimEnd().Length;
			}
		}

		protected override void OnTextChanged(EventArgs e) {
			base.OnTextChanged(e);
			UpdateRemoteContentHighlight();
		}

		private void ApplySpellCheckOption() {
			var options = (int)PInvoke.SendMessage(Handle, PInvoke.EM_GETLANGOPTIONS, IntPtr.Zero, IntPtr.Zero);
			if (_enableSpellCheck)
				options |= PInvoke.IMF_SPELLCHECKING;
			else
				options &= ~PInvoke.IMF_SPELLCHECKING;
			PInvoke.SendMessage(Handle, PInvoke.EM_SETLANGOPTIONS, IntPtr.Zero, (IntPtr)options);
		}

		private void UpdateRemoteContentHighlight() {
			bool shouldHighlight = !string.IsNullOrEmpty(_remoteContent) && _remoteContent != Text;

			if (shouldHighlight && !_remoteContentHighlighted) {
				_savedBackColor = BackColor;
				BackColor = Color.Orange;
				_remoteContentHighlighted = true;
			}
			else if (!shouldHighlight && _remoteContentHighlighted) {
				BackColor = _savedBackColor;
				_remoteContentHighlighted = false;
			}
		}

		protected override void OnMouseUp(MouseEventArgs e) {
			base.OnMouseUp(e);

			if (e.Button == MouseButtons.Left && ModifierKeys.HasFlag(Keys.Control)) {
				PInvoke.SendMessage(Handle, PInvoke.EM_SETOPTIONS, (IntPtr)PInvoke.ECOOP_AND, (IntPtr)(~PInvoke.ECO_AUTOWORDSELECTION));

				if (base.SelectionLength > 0) {
					string selected = SelectedText;
					if (selected.Length > 0 && char.IsWhiteSpace(selected[selected.Length - 1])) {
						base.SelectionLength = selected.TrimEnd().Length;
					}
				}
			}
		}

		protected override void WndProc(ref Message m) {
			if (m.Msg == PInvoke.WM_CONTEXTMENU) {
				if (_activeToolbar != null && _activeToolbar.Visible)
					_activeToolbar.Hide();

				if (_enableSpellCheck) {
					_spellCheckMenuShown = false;
					base.WndProc(ref m);
					if (_spellCheckMenuShown)
						return;
				}

				long lp = m.LParam.ToInt64();
				int x = (short)(lp & 0xFFFF);
				int y = (short)((lp >> 16) & 0xFFFF);
				Point screenPoint = (x == -1 && y == -1)
					? PointToScreen(GetPositionFromCharIndex(SelectionStart))
					: new Point(x, y);
				ShowPopupToolbar(screenPoint);
				return;
			}

			if (m.Msg == PInvoke.WM_ENTERMENULOOP)
				_spellCheckMenuShown = true;

			base.WndProc(ref m);
		}

		protected override void OnMouseDown(MouseEventArgs e) {
			if (e.Button == MouseButtons.Right) {
				int charIndex = GetCharIndexFromPosition(e.Location);
				if (charIndex < 0) charIndex = 0;
				if (charIndex > TextLength) charIndex = TextLength;

				if (SelectionLength == 0 || charIndex < SelectionStart || charIndex >= SelectionStart + SelectionLength) {
					// Use Select() to atomically set both start and length. Setting SelectionStart
					// on its own preserves the existing selection end, producing a large intermediate
					// range that can cause the control to scroll before SelectionLength = 0 collapses it.
					Select(charIndex, 0);
				}

				return;
			}

			if (e.Button == MouseButtons.Left && ModifierKeys.HasFlag(Keys.Control)) {
				PInvoke.SendMessage(Handle, PInvoke.EM_SETOPTIONS, (IntPtr)PInvoke.ECOOP_OR, (IntPtr)PInvoke.ECO_AUTOWORDSELECTION);
			}

			base.OnMouseDown(e);
		}

		private void ShowPopupToolbar(Point screenLocation) {
			if (DisableEditingPopup) return;
			int currentDpi = DeviceDpi;
			if (!_iconCaches.TryGetValue(currentDpi, out var iconCache)) {
				iconCache = new ToolbarIconCache(currentDpi / 96f);
				_iconCaches[currentDpi] = iconCache;
			}

			if (_activeToolbar == null || _activeToolbarDpi != currentDpi) {
				_activeToolbar?.Dispose();
				_activeToolbar = new PopupToolbar(this, iconCache);
				_activeToolbarDpi = currentDpi;
			}

			_activeToolbar.ShowAt(screenLocation);
		}

		#endregion

		#region - Text Formatting -

		/// <summary>
		/// Toggles bold formatting on the current selection.
		/// </summary>
		public void ToggleBold() {
			if (!EnableBold) throw new InvalidOperationException("Bold formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					font.Bold = font.Bold == tomConstants.tomTrue ? tomConstants.tomFalse : tomConstants.tomTrue;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Toggles italic formatting on the current selection.
		/// </summary>
		public void ToggleItalic() {
			if (!EnableItalic) throw new InvalidOperationException("Italic formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					font.Italic = font.Italic == tomConstants.tomTrue ? tomConstants.tomFalse : tomConstants.tomTrue;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Toggles underline formatting on the current selection.
		/// </summary>
		/// <param name="style">The underline style to apply. Defaults to <see cref="UnderlineStyle.Single"/>.</param>
		public void ToggleUnderline(UnderlineStyle style = UnderlineStyle.Single) {
			if (!EnableUnderline) throw new InvalidOperationException("Underline formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					font.Underline = font.Underline == (int)style ? tomConstants.tomNone : (int)style;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Toggles strikethrough formatting on the current selection.
		/// </summary>
		public void ToggleStrikeThrough() {
			if (!EnableStrikeThrough) throw new InvalidOperationException("Strikethrough formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					font.StrikeThrough = font.StrikeThrough == tomConstants.tomTrue ? tomConstants.tomFalse : tomConstants.tomTrue;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Sets the font color on the current selection.
		/// </summary>
		/// <param name="color">The color to apply. Use <see cref="Color.Empty"/> to reset to the automatic color.</param>
		public void SetFontColor(Color color) {
			if (!EnableFontColor) throw new InvalidOperationException("Font color formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					font.ForeColor = color.IsEmpty ? tomConstants.tomAutoColor : ColorTranslator.ToOle(color);
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Sets the background color on the current selection.
		/// </summary>
		/// <param name="color">The color to apply. Use <see cref="Color.Empty"/> to reset to the automatic color.</param>
		public void SetBackgroundColor(Color color) {
			if (!EnableBackgroundColor) throw new InvalidOperationException("Background color formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					font.BackColor = color.IsEmpty ? tomConstants.tomAutoColor : ColorTranslator.ToOle(color);
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Sets the font name on the current selection.
		/// </summary>
		/// <param name="name">The font family name to apply.</param>
		public void SetFontName(string name) {
			if (!EnableFontName) throw new InvalidOperationException("Font name formatting is disabled.");
			if (string.IsNullOrEmpty(name)) throw new ArgumentException("Font name cannot be null or empty.", nameof(name));
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					font.Name = name;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Sets the font size on the current selection.
		/// </summary>
		/// <param name="size">The font size in points. Must be greater than zero.</param>
		public void SetFontSize(float size) {
			if (!EnableFontSize) throw new InvalidOperationException("Font size formatting is disabled.");
			if (size <= 0) throw new ArgumentOutOfRangeException(nameof(size), "Font size must be greater than zero.");
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					font.Size = size;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Increases the font size on the current selection by the specified number of points.
		/// </summary>
		/// <param name="points">The number of points to increase by. Defaults to 1.</param>
		public void IncreaseFontSize(float points = 1f) {
			if (!EnableFontSize) throw new InvalidOperationException("Font size formatting is disabled.");
			if (points <= 0) throw new ArgumentOutOfRangeException(nameof(points), "Point increment must be greater than zero.");
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					font.Size += points;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Decreases the font size on the current selection by the specified number of points.
		/// The resulting size will not go below 1 point.
		/// </summary>
		/// <param name="points">The number of points to decrease by. Defaults to 1.</param>
		public void DecreaseFontSize(float points = 1f) {
			if (!EnableFontSize) throw new InvalidOperationException("Font size formatting is disabled.");
			if (points <= 0) throw new ArgumentOutOfRangeException(nameof(points), "Point decrement must be greater than zero.");
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					float newSize = font.Size - points;
					font.Size = newSize < 1f ? 1f : newSize;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Toggles superscript formatting on the current selection.
		/// Disables subscript if superscript is being enabled.
		/// </summary>
		public void ToggleSuperscript() {
			if (!EnableSuperscript) throw new InvalidOperationException("Superscript formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					if (font.Superscript == tomConstants.tomTrue) {
						font.Superscript = tomConstants.tomFalse;
					}
					else {
						//Clear subscript first: RichEdit clears both offsets when either one is cleared.
						font.Subscript = tomConstants.tomFalse;
						font.Superscript = tomConstants.tomTrue;
					}
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Toggles subscript formatting on the current selection.
		/// Disables superscript if subscript is being enabled.
		/// </summary>
		public void ToggleSubscript() {
			if (!EnableSubscript) throw new InvalidOperationException("Subscript formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var font = range.Font;
				try {
					if (font.Subscript == tomConstants.tomTrue) {
						font.Subscript = tomConstants.tomFalse;
					}
					else {
						//Clear superscript first: RichEdit clears both offsets when either one is cleared.
						font.Superscript = tomConstants.tomFalse;
						font.Subscript = tomConstants.tomTrue;
					}
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Sets the block style (heading level, paragraph, or preformatted) on the current selection's paragraph.
		/// </summary>
		/// <param name="style">The block style to apply.</param>
		public void SetBlockStyle(BlockStyle style) {
			if (!EnableHtmlFontSizing) throw new InvalidOperationException("HTML font sizing is disabled.");
			var range = GetSelectionRange();
			try {
				range.Expand(tomConstants.tomParagraph);
				int paraEnd = range.End;
				if (paraEnd > range.Start && range.GetText2(tomConstants.tomUseCRLF) != null) {
					int textEnd = paraEnd;
					var checkRange = TextDocument.Range2(paraEnd - 1, paraEnd);
					try {
						if (checkRange.Char == '\r' || checkRange.Char == '\n')
							textEnd = paraEnd - 1;
					}
					finally {
						Marshal.ReleaseComObject(checkRange);
					}
					range.SetRange(range.Start, textEnd);
				}
				var font = range.Font;
				try {
					if (style == BlockStyle.Preformatted) {
						font.Bold = tomConstants.tomFalse;
						font.Size = BlockStyleHelper.PreformattedPointSize;
						font.Name = BlockStyleHelper.PreformattedFontName;
					}
					else if (style == BlockStyle.Paragraph) {
						font.Bold = tomConstants.tomFalse;
						font.Size = BlockStyleHelper.ParagraphPointSize;
						font.Name = "";
					}
					else {
						int level = (int)style;
						font.Bold = tomConstants.tomTrue;
						font.Size = BlockStyleHelper.HeadingPointSizes[level];
						font.Name = "";
					}
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Gets the block style of the current selection's paragraph by probing its first character.
		/// </summary>
		internal BlockStyle GetBlockStyle() {
			var range = GetSelectionRange();
			try {
				range.Expand(tomConstants.tomParagraph);
				var probe = TextDocument.Range2(range.Start, range.Start + 1);
				try {
					var font = probe.Font;
					try {
						return BlockStyleHelper.GetBlockStyleFromFont(font.Size, font.Bold == tomConstants.tomTrue, font.Name);
					}
					finally {
						Marshal.ReleaseComObject(font);
					}
				}
				finally {
					Marshal.ReleaseComObject(probe);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		#endregion

		#region - Paragraph Formatting -

		/// <summary>
		/// Toggles a list on the current selection. If the selection already uses the specified
		/// list style, the list is removed. Otherwise the specified style is applied.
		/// </summary>
		/// <param name="style">The list style to apply.</param>
		public void ToggleList(ListStyle style) {
			if (!EnableLists) throw new InvalidOperationException("List formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var para = range.Para;
				try {
					int currentType = para.ListType & 0xFFFF;
					if (currentType == (int)style) {
						para.ListType = tomConstants.tomListNone;
						para.ListLevelIndex = 0;
					}
					else {
						int listType = (int)style;
						if (style != ListStyle.Bullet) {
							listType |= tomConstants.tomListPeriod;
						}

						para.ListType = listType;
						//Number from 1, as HtmlFormatter does; RichEdit otherwise starts at 0.
						para.ListStart = 1;
					}
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Per-level indentation step (in points) applied when increasing or decreasing
		/// the list nesting level. Matches the step used by <see cref="Formatting.HtmlFormatter"/>.
		/// </summary>
		private const float ListLevelIndentPoints = 18f;

		/// <summary>
		/// Increases the list nesting level of the current selection by one.
		/// Has no effect if the selection is not in a list.
		/// </summary>
		public void IncreaseListLevel() {
			if (!EnableLists) throw new InvalidOperationException("List formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var para = range.Para;
				try {
					int listType = para.ListType;
					if (listType != tomConstants.tomListNone) {
						para.ListLevelIndex++;
						// ListLevelIndex only cycles the bullet/number glyph. Paragraph indentation
						// is not derived from it, so we must shift LeftIndent explicitly to visually
						// nest the item. Re-assigning ListType (via None -> original) refreshes the
						// bullet prefix for the new level.
						para.SetIndents(para.FirstLineIndent, para.LeftIndent + ListLevelIndentPoints, para.RightIndent);
					}
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Decreases the list nesting level of the current selection by one.
		/// Has no effect if the selection is not in a list or is already at the top level.
		/// </summary>
		public void DecreaseListLevel() {
			if (!EnableLists) throw new InvalidOperationException("List formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var para = range.Para;
				try {
					int listType = para.ListType;
					if (listType != tomConstants.tomListNone && para.ListLevelIndex > 0) {
						para.ListLevelIndex--;
						// ListLevelIndex only cycles the bullet/number glyph. Paragraph indentation
						// is not derived from it, so we must shift LeftIndent explicitly to visually
						// un-nest the item. Re-assigning ListType (via None -> original) refreshes the
						// bullet prefix for the new level.
						float newLeft = para.LeftIndent - ListLevelIndentPoints;
						if (newLeft < 0f) newLeft = 0f;
						para.SetIndents(para.FirstLineIndent, newLeft, para.RightIndent);
					}
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Sets the paragraph alignment on the current selection.
		/// </summary>
		/// <param name="alignment">The alignment to apply.</param>
		public void SetAlignment(ParagraphAlignment alignment) {
			if (!EnableAlignment) throw new InvalidOperationException("Paragraph alignment formatting is disabled.");
			var range = GetSelectionRange();
			try {
				var para = range.Para;
				try {
					para.Alignment = (int)alignment;
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Increases the left indent of the current selection by the specified number of points.
		/// </summary>
		/// <param name="points">The number of points to indent by. Defaults to 36 (0.5 inch).</param>
		public void Indent(float points = 36f) {
			if (!EnableIndent) throw new InvalidOperationException("Paragraph indentation is disabled.");
			if (points <= 0) throw new ArgumentOutOfRangeException(nameof(points), "Indent amount must be greater than zero.");
			var range = GetSelectionRange();
			try {
				var para = range.Para;
				try {
					para.SetIndents(para.FirstLineIndent, para.LeftIndent + points, para.RightIndent);
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		/// <summary>
		/// Decreases the left indent of the current selection by the specified number of points.
		/// The resulting indent will not go below zero.
		/// </summary>
		/// <param name="points">The number of points to outdent by. Defaults to 36 (0.5 inch).</param>
		public void Outdent(float points = 36f) {
			if (!EnableIndent) throw new InvalidOperationException("Paragraph indentation is disabled.");
			if (points <= 0) throw new ArgumentOutOfRangeException(nameof(points), "Outdent amount must be greater than zero.");
			var range = GetSelectionRange();
			try {
				var para = range.Para;
				try {
					float newLeft = para.LeftIndent - points;
					para.SetIndents(para.FirstLineIndent, newLeft < 0f ? 0f : newLeft, para.RightIndent);
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		#endregion

		#region - Content Insertion -

		/// <summary>
		/// Inserts a hyperlink at the current selection position.
		/// </summary>
		/// <param name="url">The URL for the hyperlink. Must be a valid absolute URL.</param>
		/// <param name="displayText">The visible text for the hyperlink.</param>
		public void InsertHyperlink(string url, string displayText) {
			if (!EnableHyperlinks) throw new InvalidOperationException("Hyperlink insertion is disabled.");
			if (!Uri.TryCreate(url, UriKind.Absolute, out _))
				throw new ArgumentException("The URL must be a valid absolute URL.", nameof(url));
			if (string.IsNullOrEmpty(displayText))
				throw new ArgumentException("Display text cannot be null or empty.", nameof(displayText));

			int cpStart = SelectionStart;
			var range = GetSelectionRange();
			try {
				range.Text = displayText;
			}
			finally {
				Marshal.ReleaseComObject(range);
			}

			var linkRange = TextDocument.Range2(cpStart, cpStart + displayText.Length);
			try {
				linkRange.URL = "\"" + url + "\"";
			}
			finally {
				Marshal.ReleaseComObject(linkRange);
			}
		}

		/// <summary>
		/// Inserts an image at the current selection position. The image data is stored as-is
		/// in the document without modification.
		/// </summary>
		/// <param name="imageStream">A stream containing the image data. Supported formats are JPEG, PNG, and GIF.</param>
		/// <param name="altText">Optional alt text for the image.</param>
		public void InsertImage(Stream imageStream, string altText = "") {
			if (!EnableImages) throw new InvalidOperationException("Image insertion is disabled.");
			if (imageStream == null) throw new ArgumentNullException(nameof(imageStream));

			byte[] imageData;
			using (var ms = new MemoryStream()) {
				imageStream.CopyTo(ms);
				imageData = ms.ToArray();
			}

			using (var ms = new MemoryStream(imageData))
			using (var image = Image.FromStream(ms)) {
				GetImageFormat(image);
			}

			if (_richEditOle == null) throw new InvalidOperationException("The control handle has not been created.");
			int hr = ImageOleObject.Insert(_richEditOle, ReoConstants.REO_CP_SELECTION, imageData, altText ?? string.Empty,
				PixelsToHimetric(DefaultImageWidth), PixelsToHimetric(DefaultImageHeight));
			if (hr < 0) Marshal.ThrowExceptionForHR(hr);
		}

		/// <summary>
		/// Inserts a downsampled thumbnail
		/// wrapped in a hyperlink to the original image URL. The thumbnail is resized to
		/// <see cref="DefaultImageWidth"/> by <see cref="DefaultImageHeight"/> using high-quality
		/// bicubic interpolation.
		/// </summary>
		/// <param name="imageUrl">The URL of the original full-size image. Must be a valid absolute URL.</param>
		/// <param name="imageStream">A stream containing the original image data. Supported formats are JPEG, PNG, and GIF.</param>
		/// <param name="altText">Optional alt text for the thumbnail.</param>
		public void InsertLinkedThumbnail(string imageUrl, Stream imageStream, string altText = "") {
			if (!EnableImages) throw new InvalidOperationException("Image insertion is disabled.");
			if (!Uri.TryCreate(imageUrl, UriKind.Absolute, out _))
				throw new ArgumentException("The URL must be a valid absolute URL.", nameof(imageUrl));
			if (imageStream == null) throw new ArgumentNullException(nameof(imageStream));

			byte[] thumbnailData;
			int thumbWidth;
			int thumbHeight;
			using (var ms = new MemoryStream()) {
				imageStream.CopyTo(ms);
				ms.Position = 0;
				using (var image = Image.FromStream(ms)) {
					GetImageFormat(image); // validate the source format is supported
					thumbnailData = CreateThumbnail(image, DefaultImageWidth, DefaultImageHeight, out thumbWidth, out thumbHeight);
				}
			}

			if (_richEditOle == null) throw new InvalidOperationException("The control handle has not been created.");
			int cpStart = SelectionStart;
			int hr = ImageOleObject.Insert(_richEditOle, ReoConstants.REO_CP_SELECTION, thumbnailData, altText ?? string.Empty,
				PixelsToHimetric(thumbWidth), PixelsToHimetric(thumbHeight));
			if (hr < 0) Marshal.ThrowExceptionForHR(hr);

			var linkRange = TextDocument.Range2(cpStart, cpStart + 1);
			try {
				linkRange.URL = "\"" + imageUrl + "\"";
			}
			finally {
				Marshal.ReleaseComObject(linkRange);
			}
		}

		/// <summary>
		/// Gets the alt text for the image at the current selection position.
		/// Returns an empty string if the selection is not positioned on an image
		/// or the image has no alt text.
		/// </summary>
		public string GetImageAltText() {
			if (_richEditOle == null) return string.Empty;
			if (ImageOleObject.TryRead(_richEditOle, SelectionStart, out _, out var alt, out _, out _))
				return alt ?? string.Empty;
			return string.Empty;
		}

		private static ImageFormat GetImageFormat(Image image) {
			if (image.RawFormat.Guid == ImageFormat.Jpeg.Guid) return ImageFormat.Jpeg;
			if (image.RawFormat.Guid == ImageFormat.Png.Guid) return ImageFormat.Png;
			if (image.RawFormat.Guid == ImageFormat.Gif.Guid) return ImageFormat.Gif;
			throw new ArgumentException("Unsupported image format. Only JPEG, PNG, and GIF images are supported.");
		}

		private static byte[] CreateThumbnail(Image original, int maxWidth, int maxHeight, out int outWidth, out int outHeight) {
			// Preserve aspect ratio; never upscale. If the image already fits within the max
			// box, keep its native dimensions.
			int srcW = original.Width;
			int srcH = original.Height;
			if (srcW <= maxWidth && srcH <= maxHeight) {
				outWidth = srcW;
				outHeight = srcH;
			}
			else {
				double scale = Math.Min((double)maxWidth / srcW, (double)maxHeight / srcH);
				outWidth = Math.Max(1, (int)Math.Round(srcW * scale));
				outHeight = Math.Max(1, (int)Math.Round(srcH * scale));
			}

			// Always emit PNG regardless of source format. The RichEdit static picture cache
			// reliably renders our GDI+-decoded PNGs; other formats (notably JPEG) can come
			// back blank depending on the installed codec path.
			using (var thumbnail = new Bitmap(outWidth, outHeight, PixelFormat.Format32bppArgb))
			using (var graphics = Graphics.FromImage(thumbnail)) {
				graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
				graphics.CompositingQuality = CompositingQuality.HighQuality;
				graphics.SmoothingMode = SmoothingMode.HighQuality;
				graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
				graphics.DrawImage(original, 0, 0, outWidth, outHeight);

				using (var ms = new MemoryStream()) {
					thumbnail.Save(ms, ImageFormat.Png);
					return ms.ToArray();
				}
			}
		}

		private static int PixelsToHimetric(int pixels) {
			return (int)(pixels * 2540L / 96);
		}

		private static System.Runtime.InteropServices.ComTypes.IStream CreateComStream(byte[] data) {
			int hr = PInvoke.CreateStreamOnHGlobal(IntPtr.Zero, true, out var stream);
			if (hr != 0) Marshal.ThrowExceptionForHR(hr);
			stream.Write(data, data.Length, IntPtr.Zero);
			stream.Seek(0, 0, IntPtr.Zero);
			return stream;
		}

		#endregion

		#region - Hidden Properties -

		// Hide RTF properties

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new string? Rtf {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new string? SelectedRtf {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		// Hide text formatting properties

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new HorizontalAlignment SelectionAlignment {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new Color SelectionBackColor {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new bool SelectionBullet {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int SelectionCharOffset {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new Color SelectionColor {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new Font? SelectionFont {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int SelectionHangingIndent {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int SelectionIndent {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new bool SelectionProtected {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int SelectionRightIndent {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int[] SelectionTabs {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int BulletIndent {
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		// Hide file save/load methods

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void LoadFile(string path) {
			throw new NotSupportedException();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void LoadFile(string path, RichTextBoxStreamType fileType) {
			throw new NotSupportedException();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void LoadFile(Stream data, RichTextBoxStreamType fileType) {
			throw new NotSupportedException();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void SaveFile(string path) {
			throw new NotSupportedException();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void SaveFile(string path, RichTextBoxStreamType fileType) {
			throw new NotSupportedException();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void SaveFile(Stream data, RichTextBoxStreamType fileType) {
			throw new NotSupportedException();
		}

		#endregion
	}
}
