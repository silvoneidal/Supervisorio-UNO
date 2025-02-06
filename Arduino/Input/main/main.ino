/*
  Data: 05/01/2025
  Autor: DALÇÓQUIO AUTOMAÇÃO
  Projeto: Supervisório em Visual Basic para Arduino Uno
  Exemplo: Input - entrada digital
  Descrição: Recebe o valor do supervisório listbox Input, 
  para alterar o estado atual de uma entrada digital.
 
*/

///////////////////////////////////////////////////////////////////
// FUNÇÃO DE SETUP
void setup() {
  Serial.begin(9600);

  pinMode(2, INPUT);
  pinMode(13, OUTPUT);
  
}// end setup

///////////////////////////////////////////////////////////////////
// FUNÇÃO DE LOOP
void loop() {
    if(Serial.available()){
      String receivedData = Serial.readStringUntil('\n');
      receivedData.trim();
      if (receivedData == "02:0") {
        digitalWrite(13,LOW);
      }
      if (receivedData == "02:1") {
        digitalWrite(13, HIGH);
      }
    }

} // end loop
