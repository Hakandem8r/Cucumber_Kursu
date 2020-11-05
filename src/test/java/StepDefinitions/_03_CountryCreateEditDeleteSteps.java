package StepDefinitions;

import Pages.DialogContent;
import cucumber.api.java.en.When;

public class _03_CountryCreateEditDeleteSteps {

    DialogContent dialogContent=new DialogContent();


    @When("^User edit the \"([^\"]*)\" to \"([^\"]*)\"$")
    public void user_edit_the_to(String searchOfCountryName, String newCountryName){

        // ulkeyi arabul, ve edit e tıklat
        dialogContent.editAndDeleteFunction(searchOfCountryName, "edit");

        dialogContent.findElementAndSendKeysFunction("nameInput", newCountryName);
        dialogContent.findElementAndClickFunction("saveButton");
    }

    @When("^User delete the \"([^\"]*)\"$")
    public void user_delete_the(String nameOfCountryToDelete) {

//        dialogContent.findElementAndSendKeysFunction("searchNameInput", nameOfCountryToDelete);
//        dialogContent.findElementAndClickFunction("searchButton");

        dialogContent.editAndDeleteFunction(nameOfCountryToDelete, "delete");
        dialogContent.findElementAndClickFunction("yesButton");
    }

}