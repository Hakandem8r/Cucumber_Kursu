#Senaryo ;
#1- siteyi açınız.
#2- username ve password u gönderin ve Login butonuna tıklatın.
#3- Login olduğunuzu doğrulayın.

Feature: Login Functionality

  #username ve password girerek login doğrulama
  Scenario: Login with valid username and password

    #siteyi aç - basgara a git
    Given Navigate to basqar
    #username ve passwordu gir ve login butonuna tıklat
    When Enter username and password and click Login button
    #Login olduğunu doğrula
    Then User should login successfully



    # Gherkin dilinde 4 kelime var Given, When, And ve Then
    # Senaryonun adımlarını manalı şekilde birbirine bağlamak kullanılıyor
    # BAŞKA hiç bir önemi yok, hiç bir şeyi etkilemiyor.
    # Amaç: Teknik olmayan personelin senaryoya bakarak testin ne yapıtğını anlaması