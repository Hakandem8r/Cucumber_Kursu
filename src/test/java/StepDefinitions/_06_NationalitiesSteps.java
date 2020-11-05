package StepDefinitions;

import Pages.DialogContent;
import Pages.LeftNav;
import cucumber.api.PendingException;
import cucumber.api.java.en.And;
import cucumber.api.java.en.Then;

public class _06_NationalitiesSteps {

    LeftNav leftNav = new LeftNav();
    DialogContent dialogContent=new DialogContent();

    @And("^Navigate to Nationalities page$")
    public void navigateToNationalitiesPage() {
        leftNav.findElementAndClickFunction("setupOne");
        leftNav.findElementAndClickFunction("parameters");
        leftNav.findElementAndClickFunction("nationality");
    }


    @Then("^Create a Nationality name as \"([^\"]*)\"$")
    public void createANationalityNameAs(String name) {
        dialogContent.findElementAndClickFunction("addButton");
        dialogContent.findElementAndSendKeysFunction("nameInput", name);
        dialogContent.findElementAndClickFunction("saveButton");
    }

}