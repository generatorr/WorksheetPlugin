package de.teststory.jspwiki.worksheetplugin;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.log4j.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.wiki.WikiContext;
import org.apache.wiki.WikiEngine;
import org.apache.wiki.api.exceptions.PluginException;
import org.apache.wiki.api.exceptions.ProviderException;
import org.apache.wiki.attachment.Attachment;
import org.apache.wiki.attachment.AttachmentManager;
import org.apache.wiki.util.TextUtil;

import brushed.jspwiki.tableplugin.Table;

public class WorksheetPlugin extends Table {

	/** The parameter name for setting the src. Value is <tt>{@value}</tt>. */
	public static final String PARAM_SRC = "src";
	/** The parameter name for setting the worksheet_id. Value is <tt>{@value}</tt> */
	public static final String PARAM_WORKSHEET_ID = "sheetId";
	/** The parameter name for setting the worksheet_name. Value is <tt>{@value}</tt> */
	public static final String PARAM_WORKSHEET_NAME = "sheetName";
	/** The parameter name for showing the wiki source text instead of the table. Value is <tt>{@value}</tt> */
	public static final String PARAM_SHOW_WIKISOURCE = "showsource";

	private static final String NL = System.getProperty("line.separator");
	private static final Pattern PIPE_PATTERN = Pattern.compile("\\|");
	private static final Pattern NL_PATTERN = Pattern.compile("\r?\n");

	private final Logger log = Logger.getLogger(WorksheetPlugin.class);
	
	@SuppressWarnings({ "rawtypes", "unchecked" })
	@Override
	public String execute(WikiContext context, Map params) throws PluginException {
		log.info("WorksheetPlugin executed");
		try {
			// Analyze Parameters
			String src = getCleanParameter(params, PARAM_SRC);
			if (src == null) {
				throw new PluginException("Parameter '" + PARAM_SRC + "' is required for Workbook plugin");
			}
			int worksheetId = -1;
			String s = getCleanParameter(params, PARAM_WORKSHEET_ID);
			if (s != null) {
				try {
					worksheetId = Integer.valueOf(s).intValue();
				} catch (NumberFormatException e) {
					throw new PluginException("Parameter '" + PARAM_WORKSHEET_ID + "' must be numeric: " + s);
				}
			}
			String worksheetName = getCleanParameter(params, PARAM_WORKSHEET_NAME); // may be null
			s = getCleanParameter(params, PARAM_SHOW_WIKISOURCE);
			boolean showWikiSource = s != null ? Boolean.valueOf(s).booleanValue() : false;

			// Load the workbook
			WikiEngine engine = context.getEngine();
			AttachmentManager mgr = engine.getAttachmentManager();
			Attachment att = mgr.getAttachmentInfo(context, src);
			if (att == null) {
				throw new PluginException("Attachment '" + src + "' not found.");
			}
			InputStream inStream = mgr.getAttachmentStream(context, att);
			Workbook wb = WorkbookFactory.create(inStream);
			if (wb == null) {
				throw new PluginException("Could not load Workbook '" + src + "'");
			}

			// find worksheet
			Sheet sheet = findSheet(wb, worksheetId, worksheetName);
			if (sheet == null) {
				throw new PluginException(
						"Could not find Worksheet. Index=" + worksheetId + ", name='" + worksheetName + "'");
			}

			// Create Wiki table
			String body = analyzeSheet(sheet);

			if (showWikiSource) {
				return "<pre>" + body + "</pre>";
			} else {
				// Format the table using the "Table" Plugin
				params.put(PARAM_BODY, body);
				return super.execute(context, params);
			}
		} catch (ProviderException e) {
			throw new PluginException("Attachment info failed: " + e.getMessage(), e);
		} catch (IOException e) {
			throw new PluginException("Attachment info failed: " + e.getMessage(), e);
		} catch (EncryptedDocumentException e) {
			throw new PluginException("Attachment info failed: " + e.getMessage(), e);
		} catch (InvalidFormatException e) {
			throw new PluginException("Attachment info failed: " + e.getMessage(), e);
		}
	}

	/**
	 * This method is used to clean away things like quotation marks which a
	 * malicious user could use to stop processing and insert javascript.
	 */
	private static String getCleanParameter(Map<String, String> params, String paramId) {
		return TextUtil.replaceEntities(params.get(paramId));
	}

	/**
	 * Try to find sheet by id, then by name. Default is first sheet.
	 * 
	 * @param wb
	 * @param sheetId
	 * @param sheetName
	 * @return
	 */
	protected Sheet findSheet(Workbook wb, int sheetId, String sheetName) {
		Sheet ret = null;
		// sheet ID
		if (sheetId >= 0 && sheetId < wb.getNumberOfSheets()) {
			ret = wb.getSheetAt(sheetId);
		}
		// sheet name
		if (ret == null && sheetName != null) {
			ret = wb.getSheet(sheetName);
		}
		if (ret == null && wb.getNumberOfSheets() > 0) {
			ret = wb.getSheetAt(0);
		}
		return ret;
	}

	/**
	 * Transforms the content of the sheet into Wiki Table Syntax, as Input for
	 * the Table plugin.
	 * 
	 * @param sheet
	 * @return
	 */
	protected String analyzeSheet(Sheet sheet) {
		StringBuilder sb = new StringBuilder();

		int maxRowNum = sheet.getLastRowNum();

		// iterate through rows to find the highest column index
		int maxColumnNum = 0;
		for (int rowNum = 0; rowNum <= maxRowNum; ++rowNum) {
			Row row = sheet.getRow(rowNum);
			int lastColumnNum = row != null ? row.getLastCellNum() : 0;
			maxColumnNum = Math.max(maxColumnNum, lastColumnNum);
		}

		// parse cells and generate wiki markup
		for (int rowNum = 0; rowNum <= maxRowNum; ++rowNum) {
			Row row = sheet.getRow(rowNum);
			if (row == null) {
				continue;
			}
			for (int colNum = 0; colNum <= maxColumnNum; ++colNum) {
				Cell cell = row.getCell(colNum);
				sb.append(analyzeCell(cell));
			}
			sb.append(NL);
		}

		return sb.toString();
	}

	protected String analyzeCell(Cell cell) {
		StringBuilder ret = new StringBuilder();
		// cell delimiter
		ret.append("|");

		// TODO check if cell is within a merged region

		// formats
		String cellFormat = "";
		// TODO add css formats
		ret.append(cellFormat);

		// Text content
		String cellValue = "";
		if (cell != null) {
			int cellType = cell.getCellType();
			if (cellType == Cell.CELL_TYPE_FORMULA) {
				cellType = cell.getCachedFormulaResultType();
			}
			switch (cellType) {
			case Cell.CELL_TYPE_BLANK:
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				cellValue = cell.getBooleanCellValue() ? "Yes" : "No";
				break;
			case Cell.CELL_TYPE_ERROR:
				cellValue = "Error code: " + cell.getErrorCellValue();
				break;
			case Cell.CELL_TYPE_FORMULA:
				cellValue = cell.getCellFormula();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				// might be also a date
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					cellValue = String.valueOf(cell.getDateCellValue());
				} else {
					cellValue = String.valueOf(cell.getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cellValue = cell.getStringCellValue();
				break;
			default:
				ret.append("unknown cell type: " + cellType);
			}
		}
		ret.append(cleanCellValue(cellValue));
		return ret.toString();

	}

	/**
	 * Replace | and Newline Characters by Wiki markup
	 * 
	 * @param in
	 * @return
	 */
	protected static String cleanCellValue(String in) {
		if (in == null || in.length() == 0) {
			in = " ";
		} else {
			in = NL_PATTERN.matcher(in).replaceAll("\\\\\\\\");
			in = PIPE_PATTERN.matcher(in).replaceAll("-");
		}
		return in;
	}

}
