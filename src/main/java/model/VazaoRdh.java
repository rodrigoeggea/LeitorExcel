package model;

public class VazaoRdh {
	private Integer numeroPosto; 
	private String  nomeUsina;
	private Double  vazoes;
	
	public VazaoRdh(Integer numeroPosto, String nomeUsina, Double vazoes) {
		this.numeroPosto = numeroPosto;
		this.nomeUsina = nomeUsina;
		this.vazoes = vazoes;
	}
	public Integer getNumeroPosto() {
		return numeroPosto;
	}
	public String getNomeUsina() {
		return nomeUsina;
	}
	public Double getVazoes() {
		return vazoes;
	}
	public void setNumeroPosto(Integer numeroPosto) {
		this.numeroPosto = numeroPosto;
	}
	public void setNomeUsina(String nomeUsina) {
		this.nomeUsina = nomeUsina;
	}
	public void setVazoes(Double vazoes) {
		this.vazoes = vazoes;
	}
	@Override
	public String toString() {
		return String.format("VazaoPosto [numeroPosto=%s, nomeUsina=%s, vazoes=%s]", numeroPosto, nomeUsina, vazoes);
	}
}
