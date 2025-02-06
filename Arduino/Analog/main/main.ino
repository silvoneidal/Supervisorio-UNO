/*
  Data: 05/01/2025
  Autor: DALÇÓQUIO AUTOMAÇÃO
  Projeto: Supervisório em Visual Basic para Arduino Uno
  Exemplo: Analog - leitura analógica
  Descrição: Envia para o valor atual de uma leitura analógica,
  que será atualizada no listbox Analog do supervisório.
 
*/

///////////////////////////////////////////////////////////////////
// FUNÇÃO DE SETUP
void setup() {
  Serial.begin(9600);
  
}// end setup

///////////////////////////////////////////////////////////////////
// FUNÇÃO DE LOOP
void loop() {
    int value_analog = analogRead(A0);
    Serial.println("A0:" + String(value_analog));
    delay(1000);

} // end loop
