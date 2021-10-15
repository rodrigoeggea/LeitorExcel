package leitorexcel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Classe auxiliar para leitura e escrita de arquivo Excel utilizando o Apache POI.
 * Utiliza sempre os parâmetros coluna e linha nessa sequeência, igual as fórmulas do excel.
 * Essa classe utiliza o INDEX das colunas e linhas iniciando em 1.
 *  
 * @author Rodrigo Eggea
 */
public class LeitorExcel {
	public static enum EXTENSAO {XLS,XLSX};
	private Path caminhoArquivo;
	private Sheet planilha;
	private InputStream arquivo;
	private Workbook workbook;
	private String extensao;
	private String nomePlanilha;
	
	
	public LeitorExcel(Path caminhoArquivo) throws IOException {
		if(Files.notExists(caminhoArquivo)) { 
			throw new IOException("ERRO - Arquivo não existe: " + caminhoArquivo); 
		}
		this.caminhoArquivo = caminhoArquivo;
		this.extensao = getExtensionByStringHandling(caminhoArquivo.toString());
		this.arquivo = Files.newInputStream(caminhoArquivo, StandardOpenOption.READ);
		
		// Abre o arquivo
		if     (extensao.equalsIgnoreCase("XLS"))  { this.workbook = WorkbookFactory.create(arquivo); 	}
		else if(extensao.equalsIgnoreCase("XLSX")) { this.workbook = new XSSFWorkbook(arquivo);      }
		else throw new RuntimeException("ERRO: Arquivo não possui extensão XLS nem XLSX.");
	}
	
	private String getExtensionByStringHandling(String filename) {
		return filename.substring(filename.lastIndexOf(".") + 1 );
	}

	public LeitorExcel(InputStream inputStream, EXTENSAO extensao) throws IOException {
		this.arquivo = inputStream;
		// Abre o arquivo
		if(extensao.equals(EXTENSAO.XLS))  { this.workbook = WorkbookFactory.create(inputStream); }
		if(extensao.equals(EXTENSAO.XLSX)) { this.workbook = new XSSFWorkbook(inputStream);       }
	}
	
	public void gravar() throws Exception {
		try {
			FileOutputStream fileOut = new FileOutputStream(caminhoArquivo.toFile());
			this.workbook.write(fileOut);
			fileOut.close(); 
			this.fechar();
		} catch (Exception e) {
			throw new Exception("Não foi possível fechar a planilha! Erro: " + e.getMessage());
		}
	}

	public void gravar(Path caminhoNovoArquivo) throws Exception {
		try {
			OutputStream fileOut = new FileOutputStream(caminhoNovoArquivo.toFile());
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
	
	/* ********************** BUSCA  CÉLULA ******************************/
	
	/**
	 * Busca célular usando índice 1-Based e na ordem Coluna e Linha.
	 * @param coluna
	 * @param linha
	 * @return
	 * @throws Exception
	 */
	private Cell getCell(int coluna, int linha) throws Exception {
		verificaLimitesIndex(coluna, linha);
		
		try {
			return this.planilha.getRow(linha-1).getCell(coluna-1);
		} catch (NullPointerException npe) {
			throw new Exception("A célula indicada não está criada na aba " + this.nomePlanilha);
		}
	}
	
	private void verificaLimitesIndex(int coluna, int linha) {
		if (linha < 1) 	throw new RuntimeException("Valor inválido de index da linha="     + linha);
		if (coluna < 1) throw new RuntimeException("Valor inválido de index linha/coluna=" + coluna);
	}
	
	/* ********************** CRIA CÉLULA ******************************/
	
	/**
	 * Cria uma nova célula 1-Based na ordem coluna e linha.
	 * @param coluna
	 * @param linha
	 * @return
	 */
	private Cell createCell(int coluna, int linha) {
		verificaLimitesIndex(coluna, linha);
		
		Row row = this.planilha.getRow(linha-1);
		if (row == null) {
			row = this.planilha.createRow(linha-1);
		}
		Cell cell = row.getCell(coluna-1);
		if (cell == null) {
			cell = row.createCell(coluna-1);
		}
		return cell;
	}
	
	/* ********************* LISTAR PLANILHAS **************************/
	/**
	 * Retorna nome de todas as planilhas.
	 * @return
	 */
	public List<String> listarPlanilhas(){
		List<String> planilhas = new ArrayList<>();
		for(int tab=0; tab < workbook.getNumberOfSheets(); tab++) {
			planilhas.add(workbook.getSheetAt(tab).getSheetName());
		}
		return planilhas;
	}
	 
	/* ********************** GET BOOLEAN ******************************/
	
	public Boolean getValorCelulaBoolean(int coluna, int linha) throws Exception {
		verificaLimitesIndex(coluna, linha);
		try {
			return getCell(coluna, linha).getBooleanCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da célula não é booleano! Erro: " + ie.getMessage());
		}
	}
	
	public Boolean getValorCelulaBoolean(String celula) throws Exception {
		int coluna = CellRangeAddress.valueOf(celula).getFirstColumn()+1;
		int linha  = CellRangeAddress.valueOf(celula).getFirstRow()+1;
		return getValorCelulaBoolean(coluna, linha);
	}
	
	/* ********************** GET DOUBLE ******************************/
	
	public Double getValorCelulaDouble(int coluna, int linha) throws Exception {
		verificaLimitesIndex(coluna, linha);
		try {
			return getCell(coluna, linha).getNumericCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da célula não é numérico! Erro: " + ie.getMessage());
		}
	}
	
	public Double getValorCelulaDouble(String celula) throws Exception {
		int coluna = CellRangeAddress.valueOf(celula).getFirstColumn()+1;
		int linha  = CellRangeAddress.valueOf(celula).getFirstRow()+1;
		return getValorCelulaDouble(coluna, linha);
	}
	
	/* ********************** GET STRING ******************************/
	
	public String getValorCelulaString(int coluna, int linha) throws Exception {
		verificaLimitesIndex(coluna, linha);
		try {
			return getCell(coluna, linha).getStringCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da célula não é textual! Erro: " + ie.getMessage());
		}
	}
	
	public String getValorCelulaString(String celula) throws Exception {
		int coluna = CellRangeAddress.valueOf(celula).getFirstColumn()+1;
		int linha  = CellRangeAddress.valueOf(celula).getFirstRow()+1;
		return getValorCelulaString(coluna, linha);
	}
	
	/* ********************** GET LOCAL DATE TIME ******************************/
	
	public LocalDateTime  getValorCelulaLocalDateTime(int coluna, int linha) throws Exception {
		verificaLimitesIndex(coluna, linha);
		try {
			return getCell(coluna, linha).getLocalDateTimeCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da célula não é do tipo data! Erro: " + ie.getMessage());
		}
	}

	public LocalDateTime getValorCelulaLocalDateTime(String celula) throws Exception {
		int coluna = CellRangeAddress.valueOf(celula).getFirstColumn()+1;
		int linha  = CellRangeAddress.valueOf(celula).getFirstRow()+1;
		return getValorCelulaLocalDateTime(coluna, linha);
	}

	/* ********************** GET BOOLEAN DE UMA COLUNA ******************************/
	public List<Boolean> getValoresColunaBoolean(int colunaInicio, int linhaInicio, int linhaFim) throws Exception {
		verificaLimitesIndex(colunaInicio, linhaInicio);
		verificaLimitesIndex(colunaInicio, linhaFim);
		List<Boolean> valores = new ArrayList<>();
		
		for(int linha=linhaInicio; linha<=linhaFim; linha++) {
			valores.add(getValorCelulaBoolean(colunaInicio, linha));
		}
		return valores;
	}
	
	public List<Boolean> getValoresColunaBoolean(String celulaInicioFim) throws Exception {
		int colunaInicio = CellRangeAddress.valueOf(celulaInicioFim).getFirstColumn()+1;
		int linhaInicio  = CellRangeAddress.valueOf(celulaInicioFim).getFirstRow()+1;
		int colunaFim    = CellRangeAddress.valueOf(celulaInicioFim).getLastColumn()+1;
		int linhaFim     = CellRangeAddress.valueOf(celulaInicioFim).getLastRow()+1;
		if(colunaInicio != colunaFim) {
			throw new RuntimeException("Coluna inicial e final deve ser a mesma.");
		}
		return getValoresColunaBoolean(colunaInicio, linhaInicio, linhaFim); 
	}
	
	/* ********************** GET DOUBLE DE UMA COLUNA ******************************/
	public List<Double> getValoresColunaDouble(int colunaInicio, int linhaInicio, int linhaFim) throws Exception {
		verificaLimitesIndex(colunaInicio, linhaInicio);
		verificaLimitesIndex(colunaInicio, linhaFim);
		List<Double> valores = new ArrayList<>();
		
		for(int linha=linhaInicio; linha<=linhaFim; linha++) {
			valores.add(getValorCelulaDouble(colunaInicio, linha));
		}
		return valores;
	}
	
	public List<Double> getValoresColunaDouble(String celulaInicioFim) throws Exception {
		int colunaInicio = CellRangeAddress.valueOf(celulaInicioFim).getFirstColumn()+1;
		int linhaInicio  = CellRangeAddress.valueOf(celulaInicioFim).getFirstRow()+1;
		int colunaFim    = CellRangeAddress.valueOf(celulaInicioFim).getLastColumn()+1;
		int linhaFim     = CellRangeAddress.valueOf(celulaInicioFim).getLastRow()+1;
		if(colunaInicio != colunaFim) {
			throw new RuntimeException("getValoresColunaDouble: Coluna inicial e final deve ser a mesma.");
		}
		return getValoresColunaDouble(colunaInicio, linhaInicio, linhaFim); 
	}
	
	/* ********************** GET STRING DE UMA COLUNA ******************************/
	public List<String> getValoresColunaString(int colunaInicio, int linhaInicio, int linhaFim) throws Exception {
		verificaLimitesIndex(colunaInicio, linhaInicio);
		verificaLimitesIndex(colunaInicio, linhaFim);
		List<String> valores = new ArrayList<>();
		
		for(int linha=linhaInicio; linha<=linhaFim; linha++) {
			valores.add(getValorCelulaString(colunaInicio, linha));
		}
		return valores;
	}
	
	public List<String> getValoresColunaString(String celulaInicioFim) throws Exception {
		int colunaInicio = CellRangeAddress.valueOf(celulaInicioFim).getFirstColumn()+1;
		int linhaInicio  = CellRangeAddress.valueOf(celulaInicioFim).getFirstRow()+1;
		int colunaFim    = CellRangeAddress.valueOf(celulaInicioFim).getLastColumn()+1;
		int linhaFim     = CellRangeAddress.valueOf(celulaInicioFim).getLastRow()+1;
		if(colunaInicio != colunaFim) {
			throw new RuntimeException("Coluna inicial e final deve ser a mesma.");
		}
		return getValoresColunaString(colunaInicio, linhaInicio, linhaFim); 
	}
	
	/* ********************** GET DATETIME DE UMA COLUNA ******************************/
	public List<LocalDateTime> getValoresColunaLocalDateTime(int colunaInicio, int linhaInicio, int linhaFim) throws Exception {
		verificaLimitesIndex(colunaInicio, linhaInicio);
		verificaLimitesIndex(colunaInicio, linhaFim);
		List<LocalDateTime> valores = new ArrayList<>();
		
		for(int linha=linhaInicio; linha<=linhaFim; linha++) {
			valores.add(getValorCelulaLocalDateTime(colunaInicio, linha));
		}
		return valores;
	}
	
	public List<LocalDateTime> getValoresColunaLocalDateTime(String celulaInicioFim) throws Exception {
		int colunaInicio = CellRangeAddress.valueOf(celulaInicioFim).getFirstColumn()+1;
		int linhaInicio  = CellRangeAddress.valueOf(celulaInicioFim).getFirstRow()+1;
		int colunaFim    = CellRangeAddress.valueOf(celulaInicioFim).getLastColumn()+1;
		int linhaFim     = CellRangeAddress.valueOf(celulaInicioFim).getLastRow()+1;
		if(colunaInicio != colunaFim) {
			throw new RuntimeException("Coluna inicial e final deve ser a mesma.");
		}
		return getValoresColunaLocalDateTime(colunaInicio, linhaInicio, linhaFim); 
	}
	
	/* ********************** GET BOOLEAN DE UMA LINHA ******************************/
	public List<Boolean> getValoresLinhaBoolean(int colunaInicio, int colunaFim, int linhaFixa) throws Exception {
		verificaLimitesIndex(colunaInicio, linhaFixa);
		verificaLimitesIndex(colunaFim, linhaFixa);
		List<Boolean> valores = new ArrayList<>();
		
		for(int coluna=colunaInicio; coluna<=colunaFim; coluna++) {
			valores.add(getValorCelulaBoolean(coluna, linhaFixa));
		}
		return valores;
	}
	
	public List<Boolean> getValoresLinhaBoolean(String celulaInicioFim) throws Exception {
		int colunaInicio = CellRangeAddress.valueOf(celulaInicioFim).getFirstColumn()+1;
		int linhaInicio  = CellRangeAddress.valueOf(celulaInicioFim).getFirstRow()+1;
		int colunaFim    = CellRangeAddress.valueOf(celulaInicioFim).getLastColumn()+1;
		int linhaFim     = CellRangeAddress.valueOf(celulaInicioFim).getLastRow()+1;
		if(linhaInicio != linhaFim) {
			throw new RuntimeException("Linha inicial e final deve ser a mesma.");
		}
		return getValoresLinhaBoolean(colunaInicio, colunaFim, linhaFim); 
	}
	
	/* ********************** GET DOUBLE DE UMA LINHA ******************************/
	public List<Double> getValoresLinhaDouble(int colunaInicio, int colunaFim, int linhaFixa) throws Exception {
		verificaLimitesIndex(colunaInicio, linhaFixa);
		verificaLimitesIndex(colunaFim, linhaFixa);
		List<Double> valores = new ArrayList<>();
		
		for(int coluna=colunaInicio; coluna<=colunaFim; coluna++) {
			valores.add(getValorCelulaDouble(coluna, linhaFixa));
		}
		return valores;
	}
	
	public List<Double> getValoresLinhaDouble(String celulaInicioFim) throws Exception {
		int colunaInicio = CellRangeAddress.valueOf(celulaInicioFim).getFirstColumn()+1;
		int linhaInicio  = CellRangeAddress.valueOf(celulaInicioFim).getFirstRow()+1;
		int colunaFim    = CellRangeAddress.valueOf(celulaInicioFim).getLastColumn()+1;
		int linhaFim     = CellRangeAddress.valueOf(celulaInicioFim).getLastRow()+1;
		if(linhaInicio != linhaFim) {
			throw new RuntimeException("Linha inicial e final deve ser a mesma.");
		}
		return getValoresLinhaDouble(colunaInicio, colunaFim, linhaFim); 
	}
	
	/* ********************** GET STRING DE UMA LINHA ******************************/
	public List<String> getValoresLinhaString(int colunaInicio, int colunaFim, int linhaFixa) throws Exception {
		verificaLimitesIndex(colunaInicio, linhaFixa);
		verificaLimitesIndex(colunaFim, linhaFixa);
		List<String> valores = new ArrayList<>();
		
		for(int coluna=colunaInicio; coluna<=colunaFim; coluna++) {
			valores.add(getValorCelulaString(coluna, linhaFixa));
		}
		return valores;
	}
	
	public List<String> getValoresLinhaString(String celulaInicioFim) throws Exception {
		int colunaInicio = CellRangeAddress.valueOf(celulaInicioFim).getFirstColumn()+1;
		int linhaInicio  = CellRangeAddress.valueOf(celulaInicioFim).getFirstRow()+1;
		int colunaFim    = CellRangeAddress.valueOf(celulaInicioFim).getLastColumn()+1;
		int linhaFim     = CellRangeAddress.valueOf(celulaInicioFim).getLastRow()+1;
		if(linhaInicio != linhaFim) {
			throw new RuntimeException("Linha inicial e final deve ser a mesma.");
		}
		return getValoresLinhaString(colunaInicio, colunaFim, linhaFim); 
	}
	
	/* ********************** GET LOCAL DATE TIME DE UMA LINHA ******************************/
	public List<LocalDateTime> getValoresLinhaLocalDateTime(int colunaInicio, int colunaFim, int linhaFixa) throws Exception {
		verificaLimitesIndex(colunaInicio, linhaFixa);
		verificaLimitesIndex(colunaFim, linhaFixa);
		List<LocalDateTime> valores = new ArrayList<>();
		
		for(int coluna=colunaInicio; coluna<=colunaFim; coluna++) {
			valores.add(getValorCelulaLocalDateTime(coluna, linhaFixa));
		}
		return valores;
	}
	
	public List<LocalDateTime> getValoresLinhaLocalDateTime(String celulaInicioFim) throws Exception {
		int colunaInicio = CellRangeAddress.valueOf(celulaInicioFim).getFirstColumn()+1;
		int linhaInicio  = CellRangeAddress.valueOf(celulaInicioFim).getFirstRow()+1;
		int colunaFim    = CellRangeAddress.valueOf(celulaInicioFim).getLastColumn()+1;
		int linhaFim     = CellRangeAddress.valueOf(celulaInicioFim).getLastRow()+1;
		if(linhaInicio != linhaFim) {
			throw new RuntimeException("Linha inicial e final deve ser a mesma.");
		}
		return getValoresLinhaLocalDateTime(colunaInicio, colunaFim, linhaFim); 
	}
	
	/* ********************** GET DOUBLE RANGE ******************************/
	public Double[][] getValoresRangeDouble(int colunaInicio, int linhaInicio, int colunaFim, int linhaFim) throws Exception {
		verificaLimitesIndex(colunaInicio, linhaInicio);
		verificaLimitesIndex(colunaFim, linhaFim);

		int numeroColunas = colunaFim - colunaInicio +1;
		int numeroLinhas = linhaFim - linhaInicio +1;
		Double[][] valores = new Double[numeroLinhas][numeroColunas];
		for(int coluna=colunaInicio, i=0; coluna <= colunaFim; coluna++, i++) {
			for(int linha=linhaInicio, j=0; linha <= linhaFim; linha++, j++) {
				valores[j][i]=getValorCelulaDouble(coluna, linha);
			}
		}
		return valores;
	}
	
	public Double[][] getValoresRangeDouble(String celulaInicioFim) throws Exception {
		int colunaInicio = CellRangeAddress.valueOf(celulaInicioFim).getFirstColumn()+1;
		int linhaInicio  = CellRangeAddress.valueOf(celulaInicioFim).getFirstRow()+1;
		int colunaFim    = CellRangeAddress.valueOf(celulaInicioFim).getLastColumn()+1;
		int linhaFim     = CellRangeAddress.valueOf(celulaInicioFim).getLastRow()+1;
		return getValoresRangeDouble(colunaInicio, linhaInicio, colunaFim, linhaFim); 
	}
	
	/* ******************** SET BOOLEAN ******************************/
	
	public Cell setValorCelulaBoolean(int coluna, int linha, Boolean valor) {
		verificaLimitesIndex(coluna, linha);
		Cell cell = null;
		try {
			cell = this.getCell(coluna, linha);
			criaCelulaSeVazia(cell);
		} catch (Exception e) {
			cell = this.createCell(coluna, linha);
		}
		cell.setCellValue(valor);
		return cell;
	}
	
	public Cell setValorCelulaBoolean(String celula, Boolean valor) throws Exception {
		int coluna = CellRangeAddress.valueOf(celula).getFirstColumn()+1;
		int linha  = CellRangeAddress.valueOf(celula).getFirstRow()+1;
		return setValorCelulaBoolean(coluna, linha, valor);
	}
	
	/* ******************** SET DOUBLE ******************************/

	public Cell setValorCelulaDouble(int coluna, int linha, Double valor) {
		verificaLimitesIndex(coluna, linha);
		Cell cell = null;
		try {
			cell = this.getCell(coluna, linha);
			criaCelulaSeVazia(cell);
		} catch (Exception e) {
			cell = this.createCell(coluna, linha);
		}

		cell.setCellValue(valor);
		return cell;
	}
	
	public Cell setValorCelulaDouble(String celula, Double valor) throws Exception {
		int coluna = CellRangeAddress.valueOf(celula).getFirstColumn()+1;
		int linha  = CellRangeAddress.valueOf(celula).getFirstRow()+1;
		return setValorCelulaDouble(coluna, linha, valor);
	}
	
	/* ******************** SET STRING ******************************/
	
	public Cell setValorCelulaString(int coluna, int linha, String valor) {
		verificaLimitesIndex(coluna, linha);
		Cell cell = null;
		try {
			cell = this.getCell(coluna, linha);
			criaCelulaSeVazia(cell);
		} catch (Exception e) {
			cell = this.createCell(coluna, linha);
		}
		cell.setCellValue(valor);
		return cell;
	}
	
	public Cell setValorCelulaString(String celula, String valor) throws Exception {
		int coluna = CellRangeAddress.valueOf(celula).getFirstColumn()+1;
		int linha  = CellRangeAddress.valueOf(celula).getFirstRow()+1;
		return setValorCelulaString(coluna, linha, valor);
	}
	
	/* ******************** SET DATE ******************************/
	
	public Cell setValorCelulaDate(int coluna, int linha, Date data) {
		verificaLimitesIndex(coluna, linha);
		Cell cell = null;
		try {
			cell = this.getCell(coluna, linha);
			criaCelulaSeVazia(cell);
		} catch (Exception e) {
			cell = this.createCell(coluna, linha);
		}
		cell.setCellValue(data);
		return cell;
	}
	
	public Cell setValorCelulaDate(String celula, Date valor) throws Exception {
		int coluna = CellRangeAddress.valueOf(celula).getFirstColumn()+1;
		int linha  = CellRangeAddress.valueOf(celula).getFirstRow()+1;
		return setValorCelulaDate(coluna, linha, valor);
	}
	
	/* ****************** METODOS AUXILIARES ************************/
	private void criaCelulaSeVazia(Cell cell) {
		int linha = cell.getRowIndex();
		int coluna = cell.getColumnIndex();
		if(cell == null || cell.getCellType() == CellType.BLANK) {
			createCell(coluna, linha);
			Row row = this.planilha.getRow(linha-1);
			cell = row.createCell(coluna-1);
		}
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
	
	public Workbook getWorkbook() {
		return workbook;
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
