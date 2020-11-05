Feature: Country Create Edit Delete Functionality

  Background:
    Given Navigate to basqar
    When Enter username and password and click Login button
    Then User should login successfully
    And Navigate to country page


  Scenario: Create a country

    When Create a country name as "yeniUlk91612" and code as "yeni9Ko1612"
    Then Success message should be displayed

  Scenario: Edit a country

    When User edit the "yeniUlk91612" to "yeniUlk91512"
    Then Success message should be displayed

  Scenario: Delete a country

    When User delete the "yeniUlk91512"
    Then Success message should be displayed