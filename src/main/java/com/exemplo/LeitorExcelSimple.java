package com.exemplo;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.Date;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LeitorExcelSimple {
	private Path caminhoArquivo;
	private Sheet planilha;
	private InputStream arquivo;
	private Workbook workbook;
	private String extensao;
	private String nomePlanilha;
	
	
	public LeitorExcelSimple(Path caminhoArquivo) throws IOException {
		if(Files.notExists(caminhoArquivo)) { 
			throw new IOException("ERRO - Arquivo não existe: " + caminhoArquivo); 
		}
		this.caminhoArquivo = caminhoArquivo;
		this.extensao = getExtensionByStringHandling(caminhoArquivo.toString());
		this.arquivo = Files.newInputStream(caminhoArquivo, StandardOpenOption.READ);
		System.out.println("EXTENSION= " + extensao);
		
		// Abre o arquivo
		abrir();
	}

	private void abrir() throws EncryptedDocumentException, IOException {
		if     (extensao.equalsIgnoreCase("XLS"))  { this.workbook = WorkbookFactory.create(arquivo); 	}
		else if(extensao.equalsIgnoreCase("XLSX")) {	this.workbook = new XSSFWorkbook(arquivo);      }
		else throw new RuntimeException("ERRO: Arquivo não possui extensão XLS nem XLSX.");
	}
	
	private String getExtensionByStringHandling(String filename) {
		return filename.substring(filename.lastIndexOf(".") + 1 );
	}
	
	public void gravar() throws Exception {
		try {
			FileOutputStream fileOut = new FileOutputStream(caminhoArquivo.toFile());
			this.workbook.write(fileOut);
			fileOut.close();
			this.fechar();
		} catch (Exception e) {
			throw new Exception("Não foi possível fechar a planilha! Erro: "
					+ e.getMessage());
		}
	}

	public void gravar(Path caminhoNovoArquivo) throws Exception {
		try {
			FileOutputStream fileOut = new FileOutputStream(caminhoArquivo.toFile());
			this.workbook.write(fileOut);
			fileOut.close();
			this.fechar();
		} catch (Exception e) {
			throw new Exception("Não foi possível fechar a planilha! Erro: " + e.getMessage());
		}
	}
	
	private void fechar() throws Exception {
		try {
			this.workbook.close();
			this.arquivo.close();
		} catch (Exception e) {
			throw new Exception("Não foi possível fechar a planilha! Erro: " + e.getMessage());
		}
	} 
	
	public void abrir(String nomePlanilha) throws Exception {
		this.planilha = this.workbook.getSheet(nomePlanilha);
		this.nomePlanilha = planilha.getSheetName();
	}
	
	public void abrir(int numeroDaPlanilha) throws Exception {
		this.planilha = this.workbook.getSheetAt(numeroDaPlanilha);
		this.nomePlanilha = planilha.getSheetName();
	}
	
	private Cell getCell(int linha, int coluna) throws Exception {
		try {
			return this.planilha.getRow(linha).getCell(coluna);
		} catch (NullPointerException npe) {
			throw new Exception("A célula indicada não está criada na aba " + this.nomePlanilha);
		}
	}
	
	private Cell createCell(int linha, int coluna) {
		Row row = this.planilha.getRow(linha);
		if (row == null) {
			row = this.planilha.createRow(linha);
		}
		Cell cell = row.getCell(coluna);
		if (cell == null) {
			cell = row.createCell(coluna);
		}
		return cell;
	}
	
	public Boolean getValorCelulaBoolean(int linha, int coluna) throws Exception {
		try {
			return getCell(linha, coluna).getBooleanCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da célula não é booleano! Erro: " + ie.getMessage());
		}
	}
	
	public Double getValorCelulaDouble(int linha, int coluna) throws Exception {
		try {
			return getCell(linha, coluna).getNumericCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da célula não é numérico! Erro: " + ie.getMessage());
		}
	}
	
	public String getValorCelulaString(int linha, int coluna) throws Exception {
		try {
			return getCell(linha, coluna).getStringCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da célula não é textual! Erro: "
					+ ie.getMessage());
		}
	}
	
	public Date  getValorCelulaDate(int linha, int coluna) throws Exception {
		try {
			return getCell(linha, coluna).getDateCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da célula não é do tipo data! Erro: "
					+ ie.getMessage());
		}
	}
	
	public Cell setValorCelulaDate(int linha, int coluna, Date data) {
		Cell cell = null;
		try {
			cell = this.getCell(linha, coluna);
		} catch (Exception e) {
			cell = this.createCell(linha, coluna);
		}
		cell.setCellValue(data);
		return cell;
	}
	
	public Cell setValorCelulaBoolean(int linha, int coluna, Boolean valor) {
		Cell cell = null;
		try {
			cell = this.getCell(linha, coluna);
		} catch (Exception e) {
			cell = this.createCell(linha, coluna);
		}
		cell.setCellValue(valor);
		return cell;
	}
	
	public Cell setValorCelulaDouble(String colunaLetra, int linha, Double valor) {
		int coluna = CellReference.convertColStringToIndex(colunaLetra);
		return setValorCelulaDouble(linha-1, coluna, valor);
	}
	
	public Cell setValorCelulaDouble(int linha, int coluna, Double valor) {
		Cell cell = null;
		try {
			cell = this.getCell(linha, coluna);
		} catch (Exception e) {
			cell = this.createCell(linha, coluna);
		}
		cell.setCellValue(valor);
		return cell;
	}
	
	public Cell setValorCelulaString(int linha, int coluna, String valor) {
		Cell cell = null;
		try {
			cell = this.getCell(linha, coluna);
		} catch (Exception e) {
			cell = this.createCell(linha, coluna);
		}
		cell.setCellValue(valor);
		return cell;
	}
	
	/* getters and Setters */

	public Path getCaminhoArquivo() {
		return caminhoArquivo;
	}

	public InputStream getArquivo() {
		return arquivo;
	}

	public Sheet getPlanilha() {
		return planilha;
	}

	public String getNomePlanilha() {
		return nomePlanilha;
	}
	
	public String toString() {
		StringBuilder sb = new StringBuilder();
		for(Row row: planilha) {
			for(Cell cell: row) {
				//sb.append(String.format("%10s", cell));
				sb.append(cell + "\t");
			}
			sb.append("\n");
		}
		return sb.toString();
	}
	
}
