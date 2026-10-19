// ============================================================
//  CONFIGURAÇÃO — edite apenas este arquivo
// ============================================================
export const CONFIG = {
  // "ID do aplicativo (cliente)" do registro no Microsoft Entra (portal.azure.com).
  // Enquanto estiver como COLE_AQUI..., o app funciona em modo local (sem login/OneDrive).
  clientId: '68761df4-e86e-434d-96d1-c52ed6760781',

  // "common" aceita conta pessoal (Outlook/Hotmail/365 Pessoal) e contas corporativas.
  authority: 'https://login.microsoftonline.com/common',

  // Permissões pedidas no login. Files.ReadWrite = ler/gravar arquivos no OneDrive do próprio usuário.
  scopes: ['User.Read', 'Files.ReadWrite'],

  // Pasta criada na raiz do OneDrive de cada usuário.
  folder: 'QMC Vistorias',
};
