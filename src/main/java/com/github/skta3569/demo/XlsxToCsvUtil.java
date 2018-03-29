package com.github.skta3569.demo;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 * XLSXファイルをCSVへ変換するクラス<br>
 * SAXパースによる変換を行う（下記リンク参照）
 *
 * @see <a href="https://poi.apache.org/spreadsheet/">公式の概要</a>
 * @see <a href="https://poi.apache.org/resources/images/ss-features.png">各APIの機能サマリ</a>
 * @see <a href="https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xssf/eventusermodel/XLSX2CSV.java">公式のサンプルソース</a>
 */
public final class XlsxToCsvUtil {

	private XlsxToCsvUtil() {
	}

	/**
	 * XLSXファイルをCSVファイルへ変換する<br>
	 * XLSXファイル内に複数シートある場合にも先頭のシートのみが対象となる
	 *
	 * @param fromXlsxPath 変換元となるXLSXファイルのパス
	 * @param toCsvPath 変換先となるCSVファイルのパス（存在するパスを指定した場合は上書き）
	 */
	public static void convert(Path fromXlsxPath, Path toCsvPath) {
		System.out.println("開始　XlsxToCsvUtil#convert");

		try (OPCPackage pkg = OPCPackage.open(fromXlsxPath.toAbsolutePath().toString(), PackageAccess.READ);
				BufferedWriter bw = Files.newBufferedWriter(toCsvPath, StandardCharsets.UTF_8)) {
			ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
			XSSFReader xssfReader = new XSSFReader(pkg);
			StylesTable styleTable = xssfReader.getStylesTable();

			try (InputStream is = xssfReader.getSheetsData().next()) {
				// パース用のハンドラ生成
				ContentHandler handler = new XSSFSheetXMLHandler(
						styleTable, null, strings, new XlsxRowHandler(bw), new DataFormatter(), false);

				// パース用のXMLリーダー生成
				XMLReader sheetParser = SAXHelper.newXMLReader();
				sheetParser.setContentHandler(handler);
				System.out.println("パース開始");
				sheetParser.parse(new InputSource(is));
				System.out.println("パース終了");
			}

		} catch (InvalidOperationException | IOException | SAXException | OpenXML4JException
				| ParserConfigurationException e) {
			System.out.println("エラー　XlsxToCsvUtil#convert");
		}

		System.out.println("終了　XlsxToCsvUtil#convert");
	}

	/**
	 *
	 */
	public static class XlsxRowHandler implements SheetContentsHandler {

		private final List<String> row = new ArrayList<>();

		private final BufferedWriter bw;

		public XlsxRowHandler(BufferedWriter bw) throws IOException {
			this.bw = bw;
		}

		@Override
		public void startRow(int rowNum) {
			System.out.println(rowNum + " 行目読み込み開始");
			row.clear();
		}

		@Override
		public void cell(String cellReference, String formattedValue, XSSFComment comment) {
			row.add(formattedValue);
		}

		@Override
		public void endRow(int rowNum) {
			try {
				bw.write(String.join(",", row.stream().map(c -> "\"" + c + "\"").collect(Collectors.toList())));
				bw.newLine();
			} catch (IOException e) {
			}
		}

		@Override
		public void headerFooter(String text, boolean isHeader, String tagName) {
		}
	}

}