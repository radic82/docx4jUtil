package it.docx4j.style.custom;

import it.docx4j.style.IDocxObjectStyle;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.STVerticalJc;

public enum DocxStyleEnum implements IDocxObjectStyle {
	TITLE("000000", "28", "Calibri", true, true, false),
	TEXT_TITLE_BLACK("000000", "28", "Calibri", true, true, false),
	TEXT_TITLE_BLUE("0000FF", "28", "Calibri", true, true, false),
	TEXT_TITLE_RED("FF0000", "28", "Calibri", true, true, false),
	TEXT_SUB_TITLE_I("000000", "24", "Calibri", true, false, true),
	TEXT_SUB_TITLE("000000", "24", "Calibri", true, false, false),
	TEXT_NORMAL("444466", "10", "Calibri", false, false, false),
	TEXT_BOLD("444466", "10", "Calibri", true, false, false),
	TEXT_ITALIC("444466", "10", "Calibri", false, false, true),
	TABLE_HEADER_CELL(true, false, false, "9", "FFFFFF", "Calibri", 0, 0, 0, 0, "ccd8e1", STVerticalJc.CENTER, JcEnumeration.CENTER, true, true, true, true, true, "DBE5ED"),
	TABLE_NO_BORDER(false, false, false, "9", "000000", "Calibri", 0, 0, 0, 0, "FFFFFF", STVerticalJc.CENTER, JcEnumeration.CENTER, true, true, true, true, true, "FFFFFF"),
	TABLE_ODD_CELL(false, false, false, "8", "000000", "Calibri", 0, 0, 0, 0, "e3f1fa", STVerticalJc.TOP, JcEnumeration.LEFT, true, true, true, true, true, "DBE5ED"),
	TABLE_EVEN_CELL(false, false, false, "8", "000000", "Calibri", 0, 0, 0, 0, "FFFFFF", STVerticalJc.TOP, JcEnumeration.LEFT, true, true, true, true, true, "DBE5ED");

	DocxStyleEnum(String fontColor, String fontSize, String fontFamily, Boolean bold, Boolean underline, Boolean italic) {
		this.fontColor = fontColor;
		this.fontSize = fontSize;
		this.fontFamily = fontFamily;
		this.bold = bold;
		this.underline = underline;
		this.italic = italic;
		this.background = null;
		this.left = 0;
		this.bottom = 0;
		this.top = 0;
		this.right = 0;
		this.verticalAlignment = STVerticalJc.TOP;
		this.horizAlignment = JcEnumeration.LEFT;
		this.borderLeft = false;
		this.borderRight = false;
		this.borderTop = false;
		this.borderBottom = false;
		this.noWrap = false;
	}

	DocxStyleEnum(boolean bold, boolean italic, boolean underline, String fontSize, String fontColor, String fontFamily, int left, int bottom, int top, int right, String background, STVerticalJc verticalAlignment, JcEnumeration horizAlignment, boolean borderLeft, boolean borderRight, boolean borderTop, boolean borderBottom, boolean noWrap, String borderColor) {
		this.bold = bold;
		this.italic = italic;
		this.underline = underline;
		this.fontSize = fontSize;
		this.fontColor = fontColor;
		this.fontFamily = fontFamily;
		this.left = left;
		this.bottom = bottom;
		this.top = top;
		this.right = right;
		this.background = background;
		this.verticalAlignment = verticalAlignment;
		this.horizAlignment = horizAlignment;
		this.borderLeft = borderLeft;
		this.borderRight = borderRight;
		this.borderTop = borderTop;
		this.borderBottom = borderBottom;
		this.noWrap = noWrap;
		this.borderColor = borderColor;
	}

	private boolean bold;
	private boolean italic;
	private boolean underline;
	private String fontSize;
	private String fontColor;
	private String fontFamily;

	// cell margins
	private int left;
	private int bottom;
	private int top;
	private int right;

	private String background;
	private STVerticalJc verticalAlignment;
	private JcEnumeration horizAlignment;

	private boolean borderLeft;
	private boolean borderRight;
	private boolean borderTop;
	private boolean borderBottom;
	private boolean noWrap;
	private String borderColor;

	@Override
	public boolean isBold() {
		return bold;
	}

	@Override
	public boolean isItalic() {
		return italic;
	}

	@Override
	public boolean isUnderline() {
		return underline;
	}

	@Override
	public String getFontSize() {
		return fontSize;
	}

	@Override
	public String getFontColor() {
		return fontColor;
	}

	@Override
	public String getFontFamily() {
		return fontFamily;
	}

	@Override
	public int getLeft() {
		return left;
	}

	@Override
	public int getBottom() {
		return bottom;
	}

	@Override
	public int getTop() {
		return top;
	}

	@Override
	public int getRight() {
		return right;
	}

	@Override
	public String getBackground() {
		return background;
	}

	@Override
	public STVerticalJc getVerticalAlignment() {
		return verticalAlignment;
	}

	@Override
	public JcEnumeration getHorizAlignment() {
		return horizAlignment;
	}

	@Override
	public boolean isBorderLeft() {
		return borderLeft;
	}

	@Override
	public boolean isBorderRight() {
		return borderRight;
	}

	@Override
	public boolean isBorderTop() {
		return borderTop;
	}

	@Override
	public boolean isBorderBottom() {
		return borderBottom;
	}

	@Override
	public boolean isNoWrap() {
		return noWrap;
	}

	@Override
	public String getBorderColor() {
		return borderColor;
	}
}
