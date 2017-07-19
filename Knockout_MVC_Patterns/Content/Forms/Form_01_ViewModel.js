
// This is the *ViewModel* - JavaScript that defines the data and behavior of the UI
function QuestionnaireViewModel(state) {
    'use strict';
    var self = this;
    // Make sure we have a state object even if it is blank
    state = state || {};
    // Display Options
    self.DisplaySectionB = ko.observable(state.DisplaySectionB);
    // Section A - note that we are using OR to eliminate null or undefined values in the state
    self.Question1 = ko.observable(state.Question1 || "");
    self.Question2 = ko.observable(state.Question2 || "");
    self.Question3 = ko.observable(state.Question3 || "");
    // Section B
    self.Question4 = ko.observable(state.Question4 || "");
    self.Question5 = ko.observable(state.Question5 || "");
    self.Question6 = ko.observable(state.Question6 || "");
    // Section C
    self.Question7 = ko.observable(state.Question7 || "");
    // For a single select dropdown we need two properties, one for the list, and another for the selected value. 
    // The list property is then removed before we send the data back to the server
    self.Question8Options = ko.observableArray(['Option A', 'Option B', 'Option C', 'Option D', 'Option E', 'Option F']);
    self.Question8 = ko.observable(state.Question8 || "");
    self.Question9 = ko.observable(state.Question9 || "");
    // Section D
    self.Question10 = ko.observable(state.Question10 || "");
    self.Question11 = ko.observable(state.Question11 || "");
    // Calculations
    self.Item1 = ko.observable(state.Item1 || 0);
    self.Item2 = ko.observable(state.Item2 || 0);
    self.Item3 = ko.observable(state.Item3 || 0);
    self.Item4 = ko.observable(state.Item4 || 0);
    self.Item5 = ko.observable(state.Item5 || 0);

    // Calculated Values
    self.ItemsTotal = ko.computed(function () {
        var total = 0;
        total += parseInt(self.Item1(), 10);
        total += parseInt(self.Item2(), 10);
        total += parseInt(self.Item3(), 10);
        total += parseInt(self.Item4(), 10);
        total += parseInt(self.Item5(), 10);
        return total;
    });

    // Send final state to server to be preserved
    // Note that we use an ajax post here is that appears to work best with MVC
    self.Save = function () {
        var payload = ko.toJSON(self);
        $.ajax("/Forms/Save/", {
            type: "POST",
            processData: false,
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify({ Json: payload })
        })
        .done(function (data) {
            alert('Data saved');
        })
        .fail(function () {
            alert('Data not saved.');
        });
    };

    // Reset certain controls via a click binding
    self.Reset = function () {
        self.Item1('0');
        self.Item2('0');
        self.Item3('0');
        self.Item4('0');
        self.Item5('0');
    };

    // Copy data between observable fields using a clickbinding with a parameter
    self.CopyData = function (source, target) {
        var obsSource, obsTarget;
        for (var i = 0; i < source.length; i++) {
            //self[target[i]](self[source[i]]());
            obsSource = self[source[i]];
            obsTarget = self[target[i]];
            obsTarget(obsSource());
        };
    };
};

// Customize the stringification of the ViewModel
QuestionnaireViewModel.prototype.toJSON = function () {
    var copy = ko.toJS(this); // Get a clean copy
    delete copy.Question8Options; // Remove unneeded property
    return copy; // Return the copy to be serialized
};

// Activate knockout.js
ko.applyBindings(new QuestionnaireViewModel(state));


