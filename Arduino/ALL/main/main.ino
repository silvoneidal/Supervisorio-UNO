/*
  Data: 05/01/2025
  Autor: DALÇÓQUIO AUTOMAÇÃO
  Projeto: Supervisório em Visual Basic para Arduino Uno
  Exemplo: Input, Output, Analog e Pwm
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

      // Tratamento para Input
      if (receivedData == "02:0") {
        Serial.println("02:0"); // output
      }
      else if (receivedData == "02:1") {
        Serial.println("02:1"); // output
      }

      else{
        // Tratamento para Pwm
        int index = receivedData.indexOf(':');
        String pinStr = receivedData.substring(0, index);
        int pin = pinStr.toInt();
        String pwmStr = receivedData.substring(index + 1);
        int pwm_value = pwmStr.toInt();
        analogWrite(pin, pwm_value);
      }
    }

    // Tratamento para Analog
    int value_analog = analogRead(A0);
    Serial.println("A0:" + String(value_analog));
    delay(1000);

    // Tratamento para Output
    digitalWrite(13, HIGH);
    Serial.println("13:1");
    delay(500);

    digitalWrite(13, LOW);
    Serial.println("13:0");
    delay(500);

} // end loop
