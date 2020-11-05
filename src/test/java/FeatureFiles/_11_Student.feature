Feature: Students functionality

  Background:
    Given Navigate to basqar
    When Enter username and password and click Login button
    Then User should login successfully

  @RegressionTest
  Scenario: Create entrance exam

    And Click on the element in the left class
      | studentParent  |
      | studentsButton |

    And Click on the element in the content class
      | addButton |

    When User sending the keys in the form content class
      | firstNameInput | ismet |
      | lastNameInput  | temur |

    And Click in the new exam form
      | genderDropdown           |
      | option1                  |
      | gradeLevel               |
      | option1                  |
      | schoolDepartmentDropdown |
      | option1                  |
      | sectionDropdown          |
      | option1                  |
      | citizenshipDropdown      |
      | option1                  |
      | studentDocumentType      |
      | option1                  |

    When User sending the keys in the form content class
      | studentDocumentNumber | 1234567899 |

    And Click in the new exam form
      | representatives |
      | repAddBtn       |
      | repSelect       |
      | option1         |

    When User sending the keys in the form content class
      | repFirstName  | ahmet      |
      | repLastName   | yilmaz     |
      | repMobilPhone | 1234567899 |

    And Click in the new exam form
      | repAddBtn2 |
      | saveButton     |

    Then Success message should be displayed