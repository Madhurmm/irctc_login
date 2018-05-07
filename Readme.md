# Automate IRCTC login flow 
This scripts automates the login process on irctc website : https://www.irctc.co.in/eticketing/loginHome.jsf

IRCTC login page produces a different type of captcha every time we visit the site. In all there are 5 different kinds of captchas.

And this script can bypass 3 of them with 100% accuracy. The frequency of appearance of remaining two captchas is very low (and hence couldn't get enough handson).
So this script will work in most of the cases.

To accomplish this we make use of selenium webdriver and OCR
