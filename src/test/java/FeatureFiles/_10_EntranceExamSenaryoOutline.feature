#Bir önceki Exam Giriş Senaryosunu
#  Aşağıdaki değerler için çalıştırınız.
#
#| ExamName    | AcademicPeriodOption | GradeLevelOption | documentTypeOption |
#| Math exam   | option2              | option4          | option4            |
#| IT exam     | option1              | option1          | option1            |
#| Oracle exam | option2              | option2          | option2            |
#| Math exam   | option1              | option3          | option3            |


Feature: Entrance Exam functionality Scenario Online

  Background:
    Given Navigate to basqar
    When Enter username and password and click Login button
    Then User should login successfully

  Scenario Outline: Create entrance exam

    And Click on the element in the left class
      | entranceExamsParent |
      | entranceExamsSetup  |
      | entranceExamsChild  |

    And Click on the element in the content class
      | addButton |

    When User sending the keys in the form content class
      | nameInput | <examName> |

    And Click in the new exam form
      | academicPeriod         |
      | <academicPeriodOption> |
      | gradeLevel             |
      | <gradeLevelOption>     |
      | documentType           |
      | <documentTypeOption>   |
      | examsAddBtn            |
      | saveButton             |

    Then Success message should be displayed

    And Click on the element in the left class
      | entranceExamsParent |
      | entranceExamsSetup  |
      | entranceExamsChild  |

    When User delete the "<examName>"
    Then Success message should be displayed

    Examples:
      | examName     | academicPeriodOption | gradeLevelOption | documentTypeOption |
      | Math exam2   | option2              | option4          | option4            |
      | IT exam2     | option1              | option1          | option1            |
      | Oracle exam2 | option2              | option2          | option2            |
      | Math exam 2  | option1              | option3          | option3