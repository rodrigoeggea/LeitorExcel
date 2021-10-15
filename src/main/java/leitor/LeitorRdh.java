package leitor;

import java.io.InputStream;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

import leitorexcel.LeitorExcel;
import leitorexcel.LeitorExcel.EXTENSAO;
import model.VazaoRdh;

/**
 * Classe que realiza a leitura de arquivos RDH_xxxxx.xls 
 * (RDH - RELATÓRIO DIÁRIO DA SITUAÇÃO HIDRÁULICO-HIDROLÓGICA DAS USINAS HIDRELÉTRICAS DO SIN)
 * 
 * @author Rodrigo Eggea
 *
 */
public class LeitorRdh {
	private LeitorExcel leitor;
	private Map<Integer,VazaoRdh> postosVazoes;

	public LeitorRdh(Path path) throws Exception {
		this.leitor = new LeitorExcel(path);
		String nomePlanilha = this.leitor.listarPlanilhas().stream()
				.filter(nome -> nome.equals("Hidráulico-Hidrológica"))
				.findFirst().get();
		leitor.abrir(nomePlanilha);
		carregaVazoes();
	}

	public LeitorRdh(InputStream is, EXTENSAO ext) throws Exception {
		leitor = new LeitorExcel(is, ext);
		String nomePlanilha = this.leitor.listarPlanilhas().stream()
				.filter(nome -> nome.equals("Hidráulico-Hidrológica"))
				.findFirst().get();
		leitor.abrir(nomePlanilha);
		carregaVazoes();
	}

	private void carregaVazoes() {
		postosVazoes = new HashMap<>();
		for(int linha=8; linha <= 172; linha++) {
			try { 
				Integer numeroPosto = leitor.getValorCelulaDouble("E" + linha).intValue();
				String nomeUsina    = leitor.getValorCelulaString("C" + linha);
				Double vazao        = leitor.getValorCelulaDouble("F" + linha);
				VazaoRdh vazaoPosto = new VazaoRdh(numeroPosto, nomeUsina, vazao);
				postosVazoes.put(numeroPosto, vazaoPosto);
			} catch (Exception e) {
				// Se der erro na leitura da célula não faz nada.
			}
		}
	}
	
	public VazaoRdh getPostoVazao(int numeroPosto) {
		return postosVazoes.get(numeroPosto);
	}
	
	public Double getVazao(int numeroPosto) {
		if(postosVazoes.get(numeroPosto) != null) {
			return postosVazoes.get(numeroPosto).getVazoes();
		}
		return null;
	}
	
	public String getNomeUsina(int numeroPosto) {
		if(postosVazoes.get(numeroPosto) != null) {
			return postosVazoes.get(numeroPosto).getNomeUsina();
		}
		return null;
	}
}
