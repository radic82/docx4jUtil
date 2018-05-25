package it.docx4j.style.custom;

import it.docx4j.style.IDocxObjectStyle;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.STVerticalJc;

public class CustomNormalStyle implements IDocxObjectStyle {
	@Override
	public boolean isBold() {
		return false;
	}

	@Override
	public boolean isItalic() {
		return false;
	}

	@Override
	public boolean isUnderline() {
		return false;
	}

	@Override
	public String getFontSize() {
		return "10";
	}

	@Override
	public String getFontColor() {
		return "444466";
	}

	@Override
	public String getFontFamily() {
		return "Verdana";
	}

	@Override
	public int getLeft() {
		return 0;
	}

	@Override
	public int getBottom() {
		return 0;
	}

	@Override
	public int getTop() {
		return 0;
	}

	@Override
	public int getRight() {
		return 0;
	}

	@Override
	public String getBackground() {
		return null;
	}

	@Override
	public STVerticalJc getVerticalAlignment() {
		return null;
	}

	@Override
	public JcEnumeration getHorizAlignment() {
		return null;
	}

	@Override
	public boolean isBorderLeft() {
		return false;
	}

	@Override
	public boolean isBorderRight() {
		return false;
	}

	@Override
	public boolean isBorderTop() {
		return false;
	}

	@Override
	public boolean isBorderBottom() {
		return false;
	}

	@Override
	public boolean isNoWrap() {
		return false;
	}

	@Override
	public String getBorderColor() {
		return "CCCCCC";
	}

}
