Feature: Citizenship with senaryooutline

  Background:
    Given Navigate to basqar
    When Enter username and password and click Login button
    Then User should login successfully
    And Navigate to Citizenship page


  Scenario Outline: Citizenship create and delete with senaryo outline

    When User Create a Citizenship name as "<CitizenshipName>" short name as "<CitizenshipShortName>"
    Then Success message should be displayed

    When User Create a Citizenship name as "<CitizenshipName>" short name as "<CitizenshipShortName>"
    Then Error message should be displayed
    And Click on close button

    Then User delete the "<CitizenshipName>"
    Then Success message should be displayed


    Examples:
      | CitizenshipName | CitizenshipShortName |
      | AhmetUlkesi88   | ahmtU1829             |
      | AyseUlkesi188   | AyseU1829             |
      | KayaUlkesi188  | Kayau1829             |