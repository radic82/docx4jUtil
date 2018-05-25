package it.docx4j.item;

import it.docx4j.style.custom.DocxStyleEnum;
import it.docx4j.style.IDocxObjectStyle;

import java.util.ArrayList;
import java.util.List;

/**
 * Class to manage table (header and rows)
 */
public class DocTable extends DocItem {

	private List<DocRow> rows;

	public DocTable() {
		this(null);
	}

	public DocTable(IDocxObjectStyle style) {
		this.style = style;
		this.rows = new ArrayList<>();
	}

	/**
	 * Default Implementation
	 *
	 * @param row
	 */
	public void add(List<Object> row) {
		add(new DocRow(row));
	}

	/**
	 * @param docRow
	 */
	public void add(DocRow docRow) {
		int rowsSize = getRowSize();
		if (rowsSize == 0) {
			docRow.setStyle(DocxStyleEnum.TABLE_HEADER_CELL);
			docRow.setHeader(true);
		} else if (rowsSize % 2 == 0) {
			docRow.setStyle(DocxStyleEnum.TABLE_EVEN_CELL);
		} else {
			docRow.setStyle(DocxStyleEnum.TABLE_ODD_CELL);
		}
		rows.add(docRow);
	}

	public int getColumnSize() {
		if (rows != null && !rows.isEmpty())
			return rows.stream().findFirst().get().getRow().size();
		return 0;
	}

	public int getRowSize() {
		if (rows != null)
			return rows.size();
		return 0;
	}

	public List<DocRow> getRows() {
		return rows;
	}

	public void setRows(List<DocRow> rows) {
		this.rows = rows;
	}
}
