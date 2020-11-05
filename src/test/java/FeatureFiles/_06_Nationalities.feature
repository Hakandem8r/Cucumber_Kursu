Feature: Nationalities functionality

  Background:
    Given Navigate to basqar
    When Enter username and password and click Login button
    Then User should login successfully
    And Navigate to Nationalities page


  Scenario: Create Nationalities
    Then Create a Nationality name as "ismetish1"
    Then Success message should be displayed

  Scenario: Edit Nationalities
    Then User edit the "ismetish1" to "ismetish21"
    Then Success message should be displayed

  Scenario: Delete Nationalities
    Then User delete the "ismetish21"
    Then Success message should be displayed


  Scenario: Nationalities All functionality
    Then Create a Nationality name as "iscb8"
    Then Success message should be displayed
    Then User edit the "iscb8" to "iscb9"
    Then Success message should be displayed
    Then User delete the "iscb9"
    Then Success message should be displayed