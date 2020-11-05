Feature: DataTable Functionality

  Background:
    Given Navigate to basqar
    When Enter username and password and click Login button
    Then User should login successfully

  Scenario: Create Country
    And Click on the element in the left class
      | setupOne   |
      | parameters |
      | country    |

    And Click on the element in the content class
      | addButton |

    When User sending the keys in the dialog content class
      | nameInput | ismetCountry1 |
      | codeInput | 1231          |

    And Click on the element in the content class
      | saveButton |

    Then Success message should be displayed


  Scenario: Create a Nationalities

    And Click on the element in the left class
      | setupOne    |
      | parameters  |
      | nationality |

    And Click on the element in the content class
      | addButton |

    When User sending the keys in the dialog content class
      | nameInput | ismetNat1 |

    And Click on the element in the content class
      | saveButton |

    Then Success message should be displayed

  Scenario: Create a Fee

    And Click on the element in the left class
      | setupOne   |
      | parameters |
      | fees       |

    And Click on the element in the content class
      | addButton |

    When User sending the keys in the dialog content class
      | nameInput       | ismetFee1 |
      | codeInput       | 12321     |
      | integrationCode | 12321     |
      | priorityCode    | 12321     |

    And Click on the element in the content class
      | saveButton |

    Then Success message should be displayed




#    Senaryo Outline da verilen değerler için tüm senaryo her satır kadar tekrar çalıştırılıyordu.
#    DataTable da ise, sadece ilgili step için toplu veri gönderme imkanı oluşuyor.

