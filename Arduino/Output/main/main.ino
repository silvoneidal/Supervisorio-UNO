/*
  Data: 05/01/2025
  Autor: DALÇÓQUIO AUTOMAÇÃO
  Projeto: Supervisório em Visual Basic para Arduino Uno
  Exemplo: Output - saida digital
  Descrição: Envia para o valor atual de uma saida digital,
  que será atualizada no listbox Output do supervisório.
 
*/

///////////////////////////////////////////////////////////////////
// FUNÇÃO DE SETUP
void setup() {
  Serial.begin(9600);

  pinMode(13, OUTPUT);
  
}// end setup

///////////////////////////////////////////////////////////////////
// FUNÇÃO DE LOOP
void loop() {
    digitalWrite(13, HIGH);
    Serial.println("13:1");
    delay(1000);

    digitalWrite(13, LOW);
    Serial.println("13:0");
    delay(1000);

} // end loop
