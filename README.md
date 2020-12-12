Purpose:
======

 With the global pandemic taking place in our daily lives, contact tracing has become part of the “new normal”. We are creating a project that makes contact tracing easier with radio-frequency identification (RFID) systems. This project is important because as COVID-19 continues to be an issue, people are looking for more technological advances to promote a healthy and safe lifestyle. If our project was implemented within the university, there would be RFID card readers by the door of every classroom. Then, students would scan their Student ID card entering the classroom. The system would generate a list of all the students who went to in-person class that day. Professors would have a previously made seating chart for students so that they can easily identify who came to class and where they sat. This application would create an easier system for faculty as they then do not have to take photos of classes or take attendance.

Table of Contents
======
- [Application Setup:](#Header)
  * [Database Setup](https://github.com/willembohrer-ndsu/CSCI-488-Human-Computer-Interaction/blob/master/Examples/Database%20Setup%20Instructions.pdf)
  * [Raspberry Pi Wiring Diagram](https://github.com/willembohrer-ndsu/CSCI-488-Human-Computer-Interaction/blob/master/Examples/Raspberry_Pi_Wiring_Diagram.png)
  * [Bootstrap Setup](#Header)
    - sudo apt-get install python3-flask
    - pip install flask-bootstrap
  ---
- [Hardware:](#Header)
  * [Raspberry Pi 3](https://www.amazon.com/CanaKit-Raspberry-Complete-Starter-Kit/dp/B01C6Q2GSY/ref=pd_ybh_a_22?_encoding=UTF8&psc=1&refRID=35T7A46PN0XMSMX3S1VC)
  * [Mifare RC522](https://www.amazon.com/HiLetgo-3pcs-RFID-Kit-Raspberry/dp/B07VLDSYRW/ref=pd_ybh_a_19?_encoding=UTF8&psc=1&refRID=5QF5F4H6F3DQATS92ZB5)
  * [Wires](https://www.amazon.com/Elegoo-EL-CP-004-Multicolored-Breadboard-arduino/dp/B01EV70C78/ref=pd_ybh_a_31?_encoding=UTF8&psc=1&refRID=NXM073ZB63B2DCSFN6ZF)
  ---
- [Software:](#Header)
  * [Demonstration](https://youtu.be/UB6OK7NcDAI)
  * [App](attendance.py)
  * [Website](templates)
    - [Home Page](templates/main.html)
  * [Website Styling](static)
    - [Logo](https://github.com/willembohrer-ndsu/CSCI-488-Human-Computer-Interaction/blob/master/static/RFID%20ID%20TRACING%20Logo.PNG)
    - [CSS](static/style.css)
    - [Bootstrap](static/bootstrap-4.4.1.css)
  ---
Authors:
======
- [Willem Bohrer](https://www.linkedin.com/in/willembohrer/)
- [Nathan Marcotte](https://www.linkedin.com/in/nathanmarcotte/)
- [Ansley Schug](https://www.linkedin.com/in/ansley-schug-51b236179/)
