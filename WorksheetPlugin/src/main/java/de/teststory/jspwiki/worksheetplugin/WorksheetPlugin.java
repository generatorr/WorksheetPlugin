/*
 * Copyright 2016 Christian Froehler
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package de.teststory.jspwiki.worksheetplugin;

import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.examples.html.ToHtml;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.wiki.WikiContext;
import org.apache.wiki.WikiEngine;
import org.apache.wiki.api.exceptions.PluginException;
import org.apache.wiki.api.exceptions.ProviderException;
import org.apache.wiki.api.plugin.WikiPlugin;
import org.apache.wiki.attachment.Attachment;
import org.apache.wiki.attachment.AttachmentManager;
import org.apache.wiki.util.TextUtil;

public class WorksheetPlugin implements WikiPlugin  {

	/** The parameter name for setting the src. Value is <tt>{@value}</tt>. */
	public static final String PARAM_SRC = "src";
	/** The parameter name for setting the worksheet_id. Value is <tt>{@value}</tt> */
	public static final String PARAM_WORKSHEET_ID = "sheetId";
	/** The parameter name for setting the worksheet_name. Value is <tt>{@value}</tt> */
	public static final String PARAM_WORKSHEET_NAME = "sheetName";

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
			
			StringBuilder sb = new StringBuilder();
			ToHtml toHtml = ToHtml.create(wb, sb);
	        sb.append("<style type=\"text/css\">");
	        toHtml.printStyles();
	        sb.append("</style>");
			toHtml.printSheet(sheet);
			return sb.toString();
			
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

}
