/*
  Data: 05/01/2025
  Autor: DALÇÓQUIO AUTOMAÇÃO
  Projeto: Supervisório em Visual Basic para Arduino Uno
  Exemplo: Pwm - saida analógica
  Descrição: Recebe o valor do supervisório listbox Pwm, 
  para alterar o estado atual de saida analógica.

*/

///////////////////////////////////////////////////////////////////
// FUNÇÃO DE SETUP
void setup() {
  Serial.begin(9600);
  
}// end setup

///////////////////////////////////////////////////////////////////
// FUNÇÃO DE LOOP
void loop() {
    if (Serial.available() > 0) {
      String receivedData = Serial.readStringUntil('\n');
      receivedData.trim();
      int index = receivedData.indexOf(':');
      String pinStr = receivedData.substring(0, index);
      int pin = pinStr.toInt();
      String pwmStr = receivedData.substring(index + 1);
      int pwm_value = pwmStr.toInt();
      analogWrite(pin, pwm_value);
      
    }

} // end loop
