export function gerarNomeArquivo(prefixo: string) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  return `${prefixo}-${timestamp}.xlsx`;
}