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

if (nomeDAaba == "Investigação"){ //Compara o NOME da Aba para saber se faz a Execução


//Função que preenche a coluna Status
function DefineStatus(linha){
  var status = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //Ativa a Planilha/aba atual e faz a contagem matriz iniciar em 1 no lugar de 0
  var celula = status.getRange(linha+1,1); //Define o intervalo STATUS
  celula.setValue("ENVIADO"); //Seta valor ENVIADO no intervalo definido
}




//Função que preenche o email
function Email_Body(linha, planilha){
  var assunto = "INVESTIGAÇÃO DE NOTIFICAÇÃO DE TECNOVIGILÂNCIA";
  var email;
  var email_destinatario;
  
    //EMAIL PARA ENVIAR
    email_destinatario = "incidentes.hrn@isgh.org.br";
    email = ".";
  
  //Função do GOOGLE API para o envio de email
  MailApp.sendEmail(email_destinatario,assunto,email,{noReply:false});
}


//Função Principal
function EMAIL_FORM_INVESTIGAÇÃO(){
  
  //Define a planilha ativa
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); //CONTADOR DE MATRIZ NO 0 
  var planilha = sheet.getDataRange();
  
  //Loop com ponto de parada na LINHA VAZIA
  for (var linha = 2; planilha.getValues()[linha][0] != null; linha++){
  
  //se STATUS está vazio então chama função Email_Body() e DefineStatus()
  if ((planilha.getValues()[linha][0] == null) || (planilha.getValues()[linha][0] == "")) {
    Email_Body(linha, planilha);
    DefineStatus(linha);
   } //FIM SE
  
  } //FIM PARA
  
  
  SpreadsheetApp.flush();//Garante a execução do código ignorando possivel cache

}//Fecha Função Enviar_Email


} else {}
