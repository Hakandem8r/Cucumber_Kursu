Feature: ApachePOI Citizen Functionality

  Background:
    Given Navigate to basqar
    When Enter username and password and click Login button
    Then User should login successfully
    And Navigate to Citizenship page

  Scenario: Create and delete citizenship from excel
    When User create citizenship with ApachePOI
    Then User delete citizenship with ApachePOI
