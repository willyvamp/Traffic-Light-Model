int val;
int red1 = 10;
int red2 = 13;
int yellow1 = 9;
int yellow2 = 12;
int green1 = 8;
int green2 = 11;

void setup() {
  // put your setup code here, to run once:
  pinMode(red1, OUTPUT);
  pinMode(red2, OUTPUT);
  pinMode(yellow1, OUTPUT);
  pinMode(yellow2, OUTPUT);
  pinMode(green1, OUTPUT);
  pinMode(green2, OUTPUT);
  Serial.begin(9600);
}

void loop() {
  // put your main code here, to run repeatedly:
  while (Serial.available()){
      val = Serial.read();
    }

    // AUTO Mode
    if(val == 'a' || val == 'z') 
    {
      digitalWrite(red1, HIGH);
      digitalWrite(yellow1, LOW);
      digitalWrite(green1, LOW);
      digitalWrite(red2, LOW);
      digitalWrite(yellow2, HIGH);
      digitalWrite(green2, LOW);
      //delay(7000);
    }
    if(val == 'b')
    {
      digitalWrite(red1, LOW);
      digitalWrite(yellow1, LOW);
      digitalWrite(green1, HIGH);
      digitalWrite(red2, HIGH);
      digitalWrite(yellow2, LOW);
      digitalWrite(green2, LOW);
      //delay(3000);
    }
    if(val == 'c')
    {
      digitalWrite(red1, LOW);
      digitalWrite(yellow1, HIGH);
      digitalWrite(green1, LOW);
      digitalWrite(red2, HIGH);
      digitalWrite(yellow2, LOW);
      digitalWrite(green2, LOW);
      //delay(7000);
    }

     if(val == 'd')
    {
      digitalWrite(red1, HIGH);
      digitalWrite(yellow1, LOW);
      digitalWrite(green1, LOW);
      digitalWrite(red2, LOW);
      digitalWrite(yellow2, LOW);
      digitalWrite(green2, HIGH);
      //delay(3000);
    }
    
    // STOP
    if(val == 'z'){
      digitalWrite(red1, LOW);
      digitalWrite(red2, LOW);
      digitalWrite(yellow1, LOW);
      digitalWrite(yellow2, LOW);
      digitalWrite(green1, LOW);
      digitalWrite(green2, LOW);
    }

    // NIGHT Mode
    if(val == 'e'){
      digitalWrite(yellow1, LOW);
      digitalWrite(yellow2, LOW);
      delay(1000);
      digitalWrite(yellow1, HIGH);
      digitalWrite(yellow2, HIGH);
      delay(1000);
    }

    // MANUAL Red1-Green2
    if(val == 'g'){
       digitalWrite(red1, HIGH);
       digitalWrite(green2, HIGH);
       digitalWrite(red2, LOW);
       digitalWrite(green1, LOW);
       delay(1000);
    }
    // MANUAL Red2-Green1
    if(val == 'h'){
       digitalWrite(red2, HIGH);
       digitalWrite(green1, HIGH);
       digitalWrite(red1, LOW);
       digitalWrite(green2, LOW);
       delay(1000);
    }

}
