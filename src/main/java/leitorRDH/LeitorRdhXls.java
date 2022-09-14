package leitorRDH;

import java.io.InputStream;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

import leitorexcel.LeitorExcel;
import leitorexcel.LeitorExcel.EXTENSAO;

/**
 * Classe que realiza a leitura de arquivos RDH_xxxxx.xls (RDH - RELATÓRIO DIÁRIO DA SITUAÇÃO HIDRÁULICO-HIDROLÓGICA DAS
 * USINAS HIDRELÉTRICAS DO SIN)
 * 
 * @author Rodrigo Eggea
 *
 */
public class LeitorRdhXls {
	private LeitorExcel leitor;
	private Map<Integer, VazaoRdh> postosVazoes;

	public LeitorRdhXls(Path path) throws Exception {
		this.leitor = new LeitorExcel(path);
		String nomePlanilha = this.leitor.listarPlanilhas().stream()
				.filter(nome -> nome.equals("Hidráulico-Hidrológica")).findFirst().get();
		leitor.abrir(nomePlanilha);
		carregaVazoes();
	}

	public LeitorRdhXls(InputStream is, EXTENSAO ext) throws Exception {
		leitor = new LeitorExcel(is, ext);
		String nomePlanilha = this.leitor.listarPlanilhas().stream()
				.filter(nome -> nome.equals("Hidráulico-Hidrológica")).findFirst().get();
		leitor.abrir(nomePlanilha);
		carregaVazoes();
	}

	private void carregaVazoes() {
		postosVazoes = new HashMap<>();
		for (int linha = 8; linha <= 172; linha++) {
			try {
				Integer  numeroPosto      = leitor.getCelulaDouble("E" + linha).intValue();
				String   nomeUsina        = leitor.getCelulaString("C" + linha);
				Double   vazaoMesAnterior = leitor.getCelulaDouble("F" + linha);
				Double   vazaoMesAtual    = leitor.getCelulaDouble("H" + linha);
				VazaoRdh vazaoPosto       = new VazaoRdh(numeroPosto, nomeUsina, vazaoMesAnterior, vazaoMesAtual);
				postosVazoes.put(numeroPosto, vazaoPosto);
			} catch (Exception e) {
				// Se der erro na leitura da célula não faz nada.
			}
		}
	}

	public Double getVazaoMesAnterior(int numeroPosto) {
		if (postosVazoes.get(numeroPosto) != null) {
			return postosVazoes.get(numeroPosto).getVazaoMesAnterior();
		}
		return null;
	}
	
	public Double getVazaoMesAtual(int numeroPosto) {
		if (postosVazoes.get(numeroPosto) != null) {
			return postosVazoes.get(numeroPosto).getVazaoMesAtual();
		}
		return null;
	}

	public String getNomeUsina(int numeroPosto) {
		if (postosVazoes.get(numeroPosto) != null) {
			return postosVazoes.get(numeroPosto).getNomeUsina();
		}
		return null;
	}
	
	@Override
	public String toString() {
		return String.format("LeitorRdhXls [leitor=%s, postosVazoes=%s]", leitor, postosVazoes);
	}

	/************** INNER CLASS **************************/

	private static class VazaoRdh {
		private Integer numeroPosto;
		private String nomeUsina;
		private Double vazaoMesAnterior;
		private Double vazaoMesAtual;

		VazaoRdh(Integer numeroPosto, String nomeUsina, Double vazaoMesAnterior, Double vazaoMesAtual) {
			this.numeroPosto = numeroPosto;
			this.nomeUsina = nomeUsina;
			this.vazaoMesAnterior = vazaoMesAnterior;
			this.vazaoMesAtual    = vazaoMesAtual;
		}
		public Integer getNumeroPosto()        { return numeroPosto; }
		public String  getNomeUsina()          { return nomeUsina;   }
		public Double  getVazaoMesAnterior()   { return vazaoMesAnterior; }
		public Double  getVazaoMesAtual()      { return vazaoMesAtual;    }

		@Override
		public String toString() {
			return String.format("VazaoRdh [numeroPosto=%s, nomeUsina=%s, vazaoMesAnterior=%s, vazaoMesAtual=%s]",
					numeroPosto, nomeUsina, vazaoMesAnterior, vazaoMesAtual);
		}
	}
}