package report;

import it.docx4j.item.DocField;
import it.docx4j.item.DocImg;
import it.docx4j.item.DocLink;
import it.docx4j.item.DocRow;
import it.docx4j.item.DocTable;
import it.docx4j.item.DocText;
import it.docx4j.section.DocSectionFields;
import it.docx4j.section.DocSectionTable;
import it.docx4j.section.DocSectionText;
import it.docx4j.section.DocSectionTitle;
import it.docx4j.section.DocSimpleSection;
import it.docx4j.style.custom.DocxStyleEnum;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.toc.TocGenerator;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class CreateSectionReport {
	private final static String SRC_RESOURCES = "src/main/resources/";
	private final static String SRC_RESOURCES_IMAGE = SRC_RESOURCES + "img/";

	private List<DocSimpleSection> sections;

	@Before
	public void before() {
		sections = new ArrayList<DocSimpleSection>() {{

			add(new DocSectionTitle("test: <_CODE/> - <_TITLE/>"));

			add(new DocSectionFields("Overview"
					, new ArrayList<DocField>() {{

				add(new DocField("test Code", "test_code_value"));
				add(new DocField("test Description", new DocLink("www.google.com", "Google")));
				add(new DocField("Men", new DocText(" Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ", DocxStyleEnum.TEXT_NORMAL)));
				add(new DocField("", new DocImg(SRC_RESOURCES_IMAGE + "1.png")));
				add(new DocField("Units", "my "));
				add(new DocField("Region", new ArrayList<String>() {{
					add("region_1");
					add("region_2");
				}}));
			}}
			));

			add(new DocSectionFields("Business Needs"
					, new ArrayList<DocField>() {{
				add(new DocField("Description Summary", null));
				add(new DocField("Impact Description", null));
			}}
			));

			DocTable docTableLever = new DocTable();
			docTableLever.add(new ArrayList<Object>() {{
				add("Code");
				add("Name");
				add("t1");
				add("t2");
			}});
			for (int i = 0; i < 5; i++) {
				docTableLever.add(new ArrayList<Object>() {{
					for (int x = 0; x < 4; x++) {
						if (x == 2)
							add(new DocImg(SRC_RESOURCES_IMAGE + "1.png"));
						else
							add("??_field_" + x + "_??");
					}
				}});
			}
			add(new DocSectionTable("Test & test 2", docTableLever));

			add(new DocSectionText("WITHOUT - TITLE Lorem ipsum dolor sit amet, consectetur adipiscing elit"
					+ ", sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
					+ " Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi "
					+ "ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in volup"
					+ "tate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat no"
					+ "n proident, sunt in culpa qui officia deserunt mollit anim id est laborum."));

			add(new DocSectionText("WITH - TITLE", "Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium dolore"
					+ "mque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto "
					+ "beatae vitae dicta sunt explicabo. Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed"
					+ " quia consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt. Neque porro quisquam est, qui dolorem ipsum "
					+ "quia dolor sit amet, consectetur, adipisci velit, sed quia non numquam eius modi tempora incidunt ut labore et dolore "
					+ "magnam aliquam quaerat voluptatem. Ut enim ad minima veniam, quis nostrum exercitationem ullam corporis suscipit laboriosam, "
					+ "nisi ut aliquid ex ea commodi consequatur? Quis autem vel eum iure reprehenderit qui in ea voluptate velit esse"
					+ "uam nihil molestiae consequatur, vel illum qui dolorem eum fugiat quo voluptas nulla pariatur?"));

			//With DOC ROW
			DocTable docTableApproval = new DocTable();
			docTableApproval.add(new DocRow(new ArrayList<Object>() {{
				add("Title");
				add("By");
				add("List Selected");
				add("Replied");
				add("on");
				add("Date");
				add("Status");
			}}));
			add(new DocSectionTable("Approval History", docTableApproval));
		}};
	}

	@After
	public void after() {
		sections = null;
	}

	@Test
	public void createDefaultReport() throws Exception {
		String newFileName = "06.docx";
		File template = new File(SRC_RESOURCES, "template.docx");
		File file = new File(SRC_RESOURCES, newFileName);
		if (template.exists()) {
			if (file.exists())
				file.delete();
			FileUtils.copyFile(template, file);

			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(file);
			sections.forEach(section -> {
				if (StringUtils.isNotEmpty(section.getTitleText()))
					wordMLPackage.getMainDocumentPart().addStyledParagraphOfText(section.getTitleStyleId(), section.getTitleText());
				wordMLPackage.getMainDocumentPart().addParagraphOfText("");
				Object obj = section.retrieveObject(wordMLPackage);
				if (obj != null) {
					if (obj instanceof List) {
						wordMLPackage.getMainDocumentPart().getContent().addAll((List<P>) obj);
					} else if (obj instanceof P) {
						wordMLPackage.getMainDocumentPart().getContent().add((P) obj);
					} else if (obj instanceof Tbl) {
						wordMLPackage.getMainDocumentPart().addObject((Tbl) obj);
					}
				}
			});
			//it must be the last operation before saving
			TocGenerator tocGenerator = new TocGenerator(wordMLPackage);
			tocGenerator.generateToc(2, " TOC \\o \"1-3\" \\h \\z \\u ", false);
			wordMLPackage.getMainDocumentPart().addParagraphOfText("");
			wordMLPackage.save(file);
		}
	}
}
