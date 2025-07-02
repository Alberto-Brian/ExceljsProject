export function gerarNomeArquivo(prefixo) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  return `${prefixo}-${timestamp}.xlsx`;
}