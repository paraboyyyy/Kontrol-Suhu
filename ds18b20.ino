#include <OneWire.h>
#include <DallasTemperature.h>
OneWire pin_DS18B20(4);
DallasTemperature DS18B20(&pin_DS18B20);

void setup(void) {
  Serial.begin(9600);
  DS18B20.begin();
}

void loop(void) {
  DS18B20.requestTemperatures();
  Serial.print(DS18B20.getTempCByIndex(0));
 }