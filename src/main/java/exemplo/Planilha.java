package exemplo;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 
 * @author C046439 - Andre Miranda Pimenta
 *
 */
public class Planilha {

	/**
	 * Representa o caminho do sistema operacional onde o arquivo ser� lido.
	 */
	private String caminhoArquivo;
	/**
	 * Representa o nome da planilha que ser� aberta, a "aba" do arquivo de
	 * planihas.
	 */
	private String nomePlanilha;
	/**
	 * Represento o pr�prio arquivo que est� sendo manuseado.
	 */
	private InputStream arquivo;
	/**
	 * Representa a "aba" da planilha selecionada que est� sendo manuseada.
	 */
	private Sheet planilha;
	/**
	 * Objeto da API do POI para trabalhar com excel.
	 */
	private Workbook workbook;

	/**
	 * Construtor de classe.
	 * 
	 * @param caminhoArquivo
	 *            : deve-se passar o caminho no padr�o java para acessar o
	 *            arquivo no sistema operacional
	 */
	public Planilha(String caminhoArquivo) {
		this.caminhoArquivo = caminhoArquivo;
	}

	/**
	 * Efetua a abertura da planilha. Toda planilha para ser lida ou altera
	 * inicialmente precisa ser aberta. No momento n�o � poss�vel abrir duas
	 * abas ao mesmo tempo: deve-se abrir uma, fech�-la, e ent�o abrir a
	 * pr�xima. Processo que pode demorar.
	 * 
	 * @param nomePlanilha
	 * @throws Exception
	 */
	public void abrir(String nomePlanilha) throws Exception {
		if (this.arquivo == null) {
			try {
				this.arquivo = new FileInputStream(this.caminhoArquivo);
			} catch (Exception e) {
				throw new Exception("Arquivo " + this.arquivo
						+ " n�o encontrado! Erro: " + e.getMessage());
			}
			try {
				this.workbook = WorkbookFactory.create(this.arquivo);
			} catch (Exception e) {
				throw new Exception(
						"N�o foi poss�vel abrir o arquivo como sendo um Excel! Erro: "
								+ e.getMessage());
			}
		}
		this.planilha = this.workbook.getSheet(nomePlanilha);
	}

	/**
	 * Realiza o processo para fechar a planilha e terminar o seu uso. Esse
	 * procedimento n�o salva o arquivo caso ele tenha sido alterado.
	 * 
	 * @throws Exception
	 */
	public void fechar() throws Exception {
		try {
			this.workbook.close();
			this.arquivo.close();
		} catch (Exception e) {
			throw new Exception("N�o foi poss�vel fechar a planilha! Erro: "
					+ e.getMessage());
		}
	}

	/**
	 * Realiza o processo para fechar a planilha e terminar o seu uso. Esse
	 * procedimento SALVA o arquivo caso ele tenha sido alterado.
	 * 
	 * @throws Exception
	 */
	public void fecharEGravar() throws Exception {
		try {
			FileOutputStream fileOut = new FileOutputStream(this.caminhoArquivo);
			this.workbook.write(fileOut);
			fileOut.close();
			this.fechar();
		} catch (Exception e) {
			throw new Exception("N�o foi poss�vel fechar a planilha! Erro: "
					+ e.getMessage());
		}
	}

	/**
	 * Realiza o processo para fechar a planilha e terminar o seu uso. Esse
	 * procedimento SALVA O CONTE�DO do arquivo EM OUTRO ARQUIVO. O arquivo
	 * original permanece inalterado.
	 * 
	 * @throws Exception
	 */
	public void fecharEGravar(String caminhoNovoArquivo) throws Exception {
		try {
			FileOutputStream fileOut = new FileOutputStream(caminhoNovoArquivo);
			this.workbook.write(fileOut);
			fileOut.close();
			this.fechar();
		} catch (Exception e) {
			throw new Exception("N�o foi poss�vel fechar a planilha! Erro: "
					+ e.getMessage());
		}
	}

	/**
	 * Esse procedimento SALVA O CONTE�DO em um arquivo novo e n�o fecha o arquivo original. 
	 * O arquivo original permanece inalterado.
	 * 
	 * @throws Exception
	 */
	public void gravar(String caminhoNovoArquivo) throws Exception {
		try {
			FileOutputStream fileOut = new FileOutputStream(caminhoNovoArquivo);
			this.workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			throw new Exception("N�o foi gravar a planilha! Erro: "
					+ e.getMessage());
		}
	}

	/**
	 * M�todo interno para recuperar uma c�lula do Excel.
	 * 
	 * @param linha
	 * @param coluna
	 * @return
	 * @throws Exception
	 */
	private Cell getCell(int linha, int coluna) throws Exception {
		try {
			return this.planilha.getRow(linha).getCell(coluna);
		} catch (NullPointerException npe) {
			throw new Exception("A c�lula indicada n�o est� criada na aba " + this.nomePlanilha);
		}
	}

	/**
	 * M�todo interno para criar uma c�lula
	 * 
	 * @param linha
	 * @param coluna
	 * @return
	 */
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

	/**
	 * Recupera um valor booleano da linha e coluna. Valores iniciam com zero
	 * (0).
	 * 
	 * @param linha
	 * @param coluna
	 * @return
	 * @throws Exception
	 */
	public Boolean getValorCelulaBoolean(int linha, int coluna)
			throws Exception {
		try {
			return getCell(linha, coluna).getBooleanCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da c�lula n�o � booleano! Erro: "
					+ ie.getMessage());
		}
	}

	/**
	 * Recupera um valor Num�rico Double da linha e coluna. Valores iniciam com
	 * zero (0).
	 * 
	 * @param linha
	 * @param coluna
	 * @return
	 * @throws Exception
	 */
	public Double getValorCelulaDouble(int linha, int coluna) throws Exception {
		try {
			return getCell(linha, coluna).getNumericCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da c�lula n�o � num�rico! Erro: "
					+ ie.getMessage());
		}
	}

	/**
	 * Recupera um valor Textual (String) da linha e coluna. Valores iniciam com
	 * zero (0).
	 * 
	 * @param linha
	 * @param coluna
	 * @return
	 * @throws Exception
	 */
	public String getValorCelulaString(int linha, int coluna) throws Exception {
		try {
			return getCell(linha, coluna).getStringCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da c�lula n�o � textual! Erro: "
					+ ie.getMessage());
		}
	}
	
	
	/**
	 * Recupera um valor de Data de uma c�lula do excel.
	 * 
	 * @param linha
	 * @param coluna
	 * @return
	 * @throws Exception
	 */
	public Date  getValorCelulaDate(int linha, int coluna) throws Exception {
		try {
			return getCell(linha, coluna).getDateCellValue();
		} catch (IllegalStateException ie) {
			throw new Exception("O valor da c�lula n�o � do tipo data! Erro: "
					+ ie.getMessage());
		}
	}
	
	/**
	 * Preenche a celula com data
	 * 
	 * @param linha
	 * @param coluna
	 * @param valor
	 * @return
	 */
	public Cell setValorCelulaDate(int linha, int coluna, Date data) {
		Cell cell = null;
		try {
			cell = this.getCell(linha, coluna);
		} catch (Exception e) {
			cell = this.createCell(linha, coluna);
		}
		//cell.setCellType(CellType.BLANK);
		cell.setCellValue(data);
		return cell;
	}


	/**
	 * Preenche a celula com valor Booleno
	 * 
	 * @param linha
	 * @param coluna
	 * @param valor
	 * @return
	 */
	public Cell setValorCelulaBoolean(int linha, int coluna, Boolean valor) {
		Cell cell = null;
		try {
			cell = this.getCell(linha, coluna);
		} catch (Exception e) {
			cell = this.createCell(linha, coluna);
		}
		//cell.setCellType(Cell.CELL_TYPE_BOOLEAN);
		cell.setCellValue(valor);
		return cell;
	}

	/**
	 * Preenche a celula com valor Double
	 * 
	 * @param linha
	 * @param coluna
	 * @param valor
	 * @return
	 */
	public Cell setValorCelulaDouble(int linha, int coluna, Double valor) {
		Cell cell = null;
		try {
			cell = this.getCell(linha, coluna);
		} catch (Exception e) {
			cell = this.createCell(linha, coluna);
		}
		//cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		cell.setCellValue(valor);
		return cell;
	}

	/**
	 * Preenche a celula com valor String
	 * 
	 * @param linha
	 * @param coluna
	 * @param valor
	 * @return
	 */
	public Cell setValorCelulaString(int linha, int coluna, String valor) {
		Cell cell = null;
		try {
			cell = this.getCell(linha, coluna);
		} catch (Exception e) {
			cell = this.createCell(linha, coluna);
		}
		//cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue(valor);
		return cell;
	}

	/* getters and Setters */

	public String getCaminhoArquivo() {
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

}
