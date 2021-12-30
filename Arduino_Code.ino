int trigPin = 3;

int echoPin = 2;

int trig2=10;
int echo2=9;

long duration;

long distance_cm; 
int dc;

void setup() {
 Serial.begin (9600); 
 pinMode(trigPin, OUTPUT); 
 pinMode(echoPin, INPUT);



 pinMode(trig2,OUTPUT);
 pinMode(echo2,INPUT);
}

void loop() {
 digitalWrite(trigPin, LOW); 
 digitalWrite(trigPin, HIGH);
 digitalWrite(trigPin, LOW); 
 
  duration = pulseIn(echoPin, HIGH);
 
 
 distance_cm = duration / 29 / 2; 
 
 Serial.print(distance_cm);
 Serial.println();
 //delay(40);



 digitalWrite(trig2, LOW); 
 digitalWrite(trig2, HIGH);
 digitalWrite(trig2, LOW); 
 duration = pulseIn(echo2, HIGH);
 dc = duration / 29 / 2; 
 
 Serial.print(dc);
 Serial.println();
 
 delay(100);
}
