package StepDefinitions;

import Pages.FormContent;
import cucumber.api.DataTable;
import cucumber.api.java.en.And;
import cucumber.api.java.en.When;

import java.util.List;

public class _09_EntranceExamSteps {

    FormContent formContent=new FormContent();

    @When("^User sending the keys in the form content class$")
    public void userSendingTheKeysInTheFormContentClass(DataTable elements) {

        List< List<String> >  elementsNameAndValue = elements.asLists(String.class);

        for(int i=0;i<elementsNameAndValue.size();i++) {
            formContent.findElementAndSendKeysFunction( elementsNameAndValue.get(i).get(0),
                    elementsNameAndValue.get(i).get(1));
        }
    }

    @And("^Click in the new exam form$")
    public void clickInTheNewExamForm(DataTable elements) {

        List<String>  elementsToClick = elements.asList(String.class);

        for(int i=0;i<elementsToClick.size();i++) {
            formContent.findElementAndClickFunction(elementsToClick.get(i));
        }
    }


}