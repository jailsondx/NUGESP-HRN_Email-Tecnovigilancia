/* 
  
  ANOTAÇÕES
  
  NAVEGUE PELA PLANILHA SE GUIANDO COMO UMA MATRIZ:...
  VAR.FUNC(PARAMENTROS)[LINHA][COLUNA];
  
  ASPAS DUPLAS PARA TEXTO.
  SIMBOLO DE SOMA(+) PARA CONTATENAR.
  
*/

//Função de Obter o NOME da Aba
var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var nomeDAaba = ss.getName();

if (nomeDAaba == "Notificação1"){//Compara o NOME da Aba para saber se faz a Execução



//Função que preenche a coluna Status
function DefineStatus(linha){
  var status = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //Ativa a Planilha/aba atual e faz a contagem matriz iniciar em 1 no lugar de 0
  var celula = status.getRange(linha+1,2); //Define o intervalo STATUS
  celula.setValue("ENVIADO"); //Seta valor ENVIADO no intervalo definido
}



//Função que preenche o email
function Email_Body(linha, planilha){
  var assunto = "NOTIFICAÇÃO DE TECNOVIGILÂNCIA";
  var EQUIPAMENTO = "Equipamento";
  var email;
  var PuxaEmailDaPlanilha = planilha.getValues()[linha][0];
  var email_destinatario;
  
  if(planilha.getValues()[linha][4] == EQUIPAMENTO){
    //EMAIL PARA TIPO EQUIPAMENTO
    email_destinatario = PuxaEmailDaPlanilha + ",walder.costa@isgh.org.br,incidentes.hrn@isgh.org.br,clarissatomas@gmail.com";
    email = "Setor: " + planilha.getValues()[linha][3] + 
      "\nEquipamento Notificado: " + planilha.getValues()[linha][5] + 
      "\nModelo e Empresa fabricante do equipamento notificado: " + planilha.getValues()[linha][6] + 
      "\nNº do patrimônio do equipamento: " + planilha.getValues()[linha][7] +
      "\nNº da tag do equipamento: " + planilha.getValues()[linha][8] + 
      "\nDescreva detalhadamente o problema apresentado pelo equipamento e as consequências para o paciente: " + planilha.getValues()[linha][9] +
      "\nNome do notificador: " + planilha.getValues()[linha][16] + 
      "\nCargo/função do notificador: " + planilha.getValues()[linha][17];
  
  } else {
    //EMAIL SEM TIPO EQUIPAMENTO
    email_destinatario = PuxaEmailDaPlanilha + ",incidentes.hrn@isgh.org.br,clarissatomas@gmail.com";
    email =  "Setor: " + planilha.getValues()[linha][3] + 
      "\nTipo de produto notificado: " + planilha.getValues()[linha][10] +
      "\nNome do produto notificado: " + planilha.getValues()[linha][11] + 
      "\nFabricante do produto notificado: " + planilha.getValues()[linha][12] +
      "\nNº do lote do produto notificado: " + planilha.getValues()[linha][13] + 
      "\nNº de registro no MS/ANVISA do produto notificado: " + planilha.getValues()[linha][14] + 
      "\nDescreva detalhadamente o problema apresentado pelo produto e as consequências para o paciente: " + planilha.getValues()[linha][15] +
      "\nNome do notificador: " + planilha.getValues()[linha][16] + 
      "\nCargo/função do notificador: " + planilha.getValues()[linha][17];
    }
  
  //Função do GOOGLE API para o envio de email
  MailApp.sendEmail(email_destinatario,assunto,email,{noReply:false});
}


//Função Principal
function EMAIL_FORM_NOTIFICAÇÃO(){
  
  //Define a planilha ativa
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //CONTADOR DE MATRIZ NO 0 
  var planilha = sheet.getDataRange();
  
  //Loop com ponto de parada na LINHA VAZIA
  for (var linha = 1; planilha.getValues()[linha][0] != null; linha++){
  
  //se STATUS está vazio então chama função Email_Body() e DefineStatus()
  if ((planilha.getValues()[linha][1] == null) || (planilha.getValues()[linha][1] == "")) {
    Email_Body(linha, planilha);
    DefineStatus(linha);
   } //FIM SE
  
  } //FIM PARA
  
  
  SpreadsheetApp.flush();//Garante a execução do código ignorando possivel cache

}//Fecha Função Enviar_Email

} else {}
