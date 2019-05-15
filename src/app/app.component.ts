import { Component } from '@angular/core';

const template = require('./app.component.html');

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent {
    welcomeMessage = 'Welcome';

    async inputChanged(event) {
        return Word.run(async context => {
            /**
             * Insert your Word code here
             */
            const thisDocument = context.document;
            const stateControls = thisDocument.contentControls.getByTag(event.target.name);
            const totalIncome = thisDocument.contentControls.getByTag("totalIncome").getFirst();
            const totalDeduction = thisDocument.contentControls.getByTag("totalDeductions").getFirst();
            const totalNetIncome = thisDocument.contentControls.getByTag("totalNetIncome").getFirst();
            const totalMonthlyExpenses = thisDocument.contentControls.getByTag("totalMonthlyExpenses").getFirst();
            const cashHandValue = thisDocument.contentControls.getByTag("cashHand").getFirst();
            const cashBankValue = thisDocument.contentControls.getByTag("cashBank").getFirst();
            const stocksValue = thisDocument.contentControls.getByTag("stocks").getFirst();
            const jewelryValue = thisDocument.contentControls.getByTag("jewelry").getFirst();
            const notesValue = thisDocument.contentControls.getByTag("notesAsset").getFirst();
            const contentsHomeValue = thisDocument.contentControls.getByTag("contentsHome").getFirst();
            const carsValue = thisDocument.contentControls.getByTag("cars").getFirst();
            const lifeValue = thisDocument.contentControls.getByTag("life").getFirst();
            const realEstateValue = thisDocument.contentControls.getByTag("realEstate").getFirst();
            const estateHomeValue = thisDocument.contentControls.getByTag("estateHome").getFirst();
            const otherAssetValue = thisDocument.contentControls.getByTag("otherAsset").getFirst();

            const totalAssetsValue = thisDocument.contentControls.getByTag("totalAssets").getFirst();
            const assetsPetitionerValue = thisDocument.contentControls.getByTag("assetsPetitioner").getFirst();
            const assetsRespondentValue = thisDocument.contentControls.getByTag("assetsRespondent").getFirst();

            const securityValue = thisDocument.contentControls.getByTag("security").getFirst();
            const creditorValue = thisDocument.contentControls.getByTag("creditor").getFirst();
            const balanceValue = thisDocument.contentControls.getByTag("balance").getFirst();
            const totalLiabilitiesValue = thisDocument.contentControls.getByTag("totalLiabilities").getFirst();
    

            stateControls.load("items");
            context.sync().then( () => {
                stateControls.items.forEach(thisControl => {
                    thisControl.insertText(event.target.value, Word.InsertLocation.replace);
                });

                var businessIncome = document.getElementById("businessIncome").value;
                var disability = document.getElementById("disability").value;
                var workerComp = document.getElementById("workerComp").value;
                var unemploymentComp = document.getElementById("unemploymentComp").value;
                var rental = document.getElementById("rental").value;
                var recurring = document.getElementById("recurring").value;
                var reimbursed = document.getElementById("reimbursed").value;
                var spousal = document.getElementById("spousal").value;
                var interest = document.getElementById("interest").value;
                var gains = document.getElementById("gains").value;
                var pension = document.getElementById("pension").value;
                var ssbenefits = document.getElementById("ssbenefits").value;
                // deduction
                var fica = document.getElementById("fica").value;
                var healthInsurance = document.getElementById("healthInsurance").value;
                var union = document.getElementById("union").value;
                var federalTaxes = document.getElementById("federalTaxes").value;
                var retirement = document.getElementById("retirement").value;

                // expenses Values
                var mortgage = document.getElementById("mortgage").value;
                var cosmetics = document.getElementById("cosmetics").value;
                var barber = document.getElementById("barber").value;
                var fuel = document.getElementById("fuel").value;
                var electricity = document.getElementById("electricity").value;
                var waste = document.getElementById("waste").value;
                var telephone = document.getElementById("telephone").value;
                var holiday = document.getElementById("holiday").value;
                var other = document.getElementById("other").value;
                var property = document.getElementById("property").value;
                var insuranceExpense = document.getElementById("insuranceExpense").value;
                // assetts
                var cashHandPetitioner = document.getElementById("cashHandPetitioner").value;
                var cashHandRespondent = document.getElementById("cashHandRespondent").value;

                var cashBankPetitioner = document.getElementById("cashBankPetitioner").value;
                var cashBankRespondent = document.getElementById("cashBankRespondent").value;
                var stocksPetitioner = document.getElementById("stocksPetitioner").value;
                var stocksRespondent = document.getElementById("stocksRespondent").value;
                var notesPetitioner = document.getElementById("notesPetitioner").value;
                var notesRespondent = document.getElementById("notesRespondent").value;
                var realEstatePetitioner = document.getElementById("realEstatePetitioner").value;
                var realEstateRespondent = document.getElementById("realEstateRespondent").value;
                var estateHomePetitioner = document.getElementById("estateHomePetitioner").value;
                var estateHomeRespondent = document.getElementById("estateHomeRespondent").value;
                var carsPetitioner = document.getElementById("carsPetitioner").value;
                var carsRespondent = document.getElementById("carsRespondent").value;
                var contentsHomePetitioner = document.getElementById("contentsHomePetitioner").value;
                var contentsHomeRespondent = document.getElementById("contentsHomeRespondent").value;
                var jewelryPetitioner = document.getElementById("jewelryPetitioner").value;
                var jewelryRespondent = document.getElementById("jewelryRespondent").value;
                var lifePetitioner = document.getElementById("lifePetitioner").value;
                var lifeRespondent = document.getElementById("lifeRespondent").value;
                var otherAssetPetitioner = document.getElementById("otherAssetPetitioner").value;
                var otherAssetRespondent = document.getElementById("otherAssetRespondent").value;
                var creditorRespondent = document.getElementById("creditorRespondent").value;
                var securityRespondent = document.getElementById("securityRespondent").value;
                var creditorPetitioner = document.getElementById("creditorPetitioner").value;
                var securityPetitioner = document.getElementById("securityPetitioner").value;

                // Calculations
                let sum = parseInt(businessIncome) + parseInt(pension) + parseInt(workerComp) +
                parseInt(disability) + parseInt(unemploymentComp) + parseInt(rental) + parseInt(recurring) + parseInt(reimbursed) +
                parseInt(spousal) + parseInt(gains) +parseInt(ssbenefits) + parseInt(interest);

                // total Deductions
                let deductions = parseInt(fica) + parseInt(retirement) + parseInt(union) +parseInt(federalTaxes) +
                parseInt(healthInsurance);

                //  total Net Income
                let net = sum - deductions;

                //  total Monthly Expenses
                let expenses = 
                parseInt(property) +
                parseInt(mortgage) + 
                parseInt(cosmetics) + 
                parseInt(barber)+ 
                parseInt(fuel) + 
                parseInt(electricity) + 
                parseInt(waste) + 
                parseInt(telephone) + 
                parseInt(holiday) + 
                parseInt(other) + 
                parseInt(insuranceExpense);

                // cashHand
                let cashHand = this.addTwoValues(cashHandPetitioner, cashHandRespondent);
                let cashBank = this.addTwoValues(cashBankPetitioner, cashBankRespondent);
                let stocks = this.addTwoValues(stocksPetitioner, stocksRespondent);
                let jewelry = this.addTwoValues(jewelryPetitioner, jewelryRespondent);
                let notes = this.addTwoValues(notesPetitioner, notesRespondent);
                let contentsHome = this.addTwoValues(contentsHomePetitioner, contentsHomeRespondent);
                let realEstate = this.addTwoValues(realEstatePetitioner, realEstateRespondent);
                let estateHome = this.addTwoValues(estateHomePetitioner, estateHomeRespondent);
                let cars = this.addTwoValues(carsPetitioner, carsRespondent);
                let life = this.addTwoValues(lifePetitioner, lifeRespondent);
                let otherAsset = this.addTwoValues(otherAssetPetitioner, otherAssetRespondent);
                // liabilities
                let security = this.addTwoValues(securityPetitioner, securityRespondent);
                let creditorDebt = this.addTwoValues(creditorPetitioner, creditorRespondent);

                   // totalAssets
                let totalAssets = cashHand + cashBank + stocks + notes + realEstate
                + estateHome + cars + jewelry + contentsHome + life + otherAsset;

                // total Petitioner
                let assetsPetitioner = parseInt(cashHandPetitioner) + parseInt(cashBankPetitioner)+ parseInt(stocksPetitioner) + parseInt(notesPetitioner) + parseInt(realEstatePetitioner) + parseInt(estateHomePetitioner) + parseInt(carsPetitioner) + parseInt(jewelryPetitioner) + parseInt(contentsHomePetitioner) + parseInt(lifePetitioner) + parseInt(otherAssetPetitioner);

                // total Respondent
                let assetsRespondent = parseInt(cashHandRespondent) + parseInt(cashBankRespondent)+ parseInt(stocksRespondent) + parseInt(notesRespondent) + parseInt(realEstateRespondent) + parseInt(estateHomeRespondent) + parseInt(carsRespondent) + parseInt(jewelryRespondent) + parseInt(contentsHomeRespondent) + parseInt(lifeRespondent) + parseInt(otherAssetRespondent);
 
                let balance = security + creditorDebt;


                // insert into word
                totalIncome.insertText(sum.toString(), Word.InsertLocation.replace);
                totalDeduction.insertText(deductions.toString(), Word.InsertLocation.replace);
                totalNetIncome.insertText(net.toString(), Word.InsertLocation.replace);
                totalMonthlyExpenses.insertText(expenses.toString(), Word.InsertLocation.replace);

                cashHandValue.insertText(cashHand.toString(), Word.InsertLocation.replace);
                cashBankValue.insertText(cashBank.toString(), Word.InsertLocation.replace);
                stocksValue.insertText(stocks.toString(), Word.InsertLocation.replace);
                jewelryValue.insertText(jewelry.toString(), Word.InsertLocation.replace);
                notesValue.insertText(notes.toString(), Word.InsertLocation.replace);
                contentsHomeValue.insertText(contentsHome.toString(), Word.InsertLocation.replace);
                lifeValue.insertText(life.toString(), Word.InsertLocation.replace)
                carsValue.insertText(cars.toString(), Word.InsertLocation.replace);
                realEstateValue.insertText(realEstate.toString(), Word.InsertLocation.replace);
                estateHomeValue.insertText(estateHome.toString(), Word.InsertLocation.replace);
                otherAssetValue.insertText(otherAsset.toString(), Word.InsertLocation.replace);

                totalAssetsValue.insertText(totalAssets.toString(), Word.InsertLocation.replace);
                assetsPetitionerValue.insertText(assetsPetitioner.toString(), Word.InsertLocation.replace);
                assetsRespondentValue.insertText(assetsRespondent.toString(), Word.InsertLocation.replace);

                creditorValue.insertText(creditorDebt.toString(), Word.InsertLocation.replace);
                securityValue.insertText(security.toString(), Word.InsertLocation.replace);
                balanceValue.insertText(balance.toString(), Word.InsertLocation.replace);
                totalLiabilitiesValue.insertText(balance.toString(), Word.InsertLocation.replace);

                return context.sync();
            });

            await context.sync();
        });
    }

    addTwoValues(value1, value2) {
        let sum = parseInt(value1) + parseInt(value2);
        return sum;
    }

    CreateNewDocument(event) {
        Word.run(async context => {
          try {
              var newDoc = context.application.createDocument('');
              newDoc.open();
              context.sync();
          } catch (error) {
              console.log(error);
          }
          event.completed();
          context.sync();
        });
    }
}




