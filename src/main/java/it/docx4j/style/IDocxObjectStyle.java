package it.docx4j.style;

import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.STVerticalJc;

import java.io.Serializable;

public interface IDocxObjectStyle extends Serializable {
	boolean isBold();

	boolean isItalic();

	boolean isUnderline();

	String getFontSize();

	String getFontColor();

	String getFontFamily();

	int getLeft();

	int getBottom();

	int getTop();

	int getRight();

	String getBackground();

	STVerticalJc getVerticalAlignment();

	JcEnumeration getHorizAlignment();

	boolean isBorderLeft();

	boolean isBorderRight();

	boolean isBorderTop();

	boolean isBorderBottom();

	boolean isNoWrap();

	String getBorderColor();
}
