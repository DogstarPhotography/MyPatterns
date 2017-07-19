
<div class="container">

    <h3>Tasks</h3>

    <form data-bind="submit: addTask">
        Add task: <input data-bind="value: newTaskText" placeholder="What needs to be done?" />
        <button type="submit">Add</button>
    </form>

    <ul class="list-group" data-bind="foreach: tasks, visible: tasks().length > 0">
        <li class="list-group-item">
            <div class="form-inline form-group">
                <input class="form-control" type="checkbox" data-bind="checked: isDone" />
                <input class="form-control" type="text" style="width: 80%;" data-bind="value: title, disable: isDone" />
                <a href="#" data-bind="click: $parent.removeTask">Delete</a>
            </div>
        </li>
    </ul>
    You have <b data-bind="text: incompleteTasks().length">&nbsp;</b> incomplete task(s)
    <span data-bind="visible: incompleteTasks().length == 0"> - it's beer time!</span>

    <button data-bind="click: save">Save</button>

</div>

@Section Scripts

    <script>
        // Define a task class
        function Task(data) {
            this.title = ko.observable(data.Title);
            this.isDone = ko.observable(data.IsDone);
        }

        // This version of a view model fetches data from the server when it is initialized and posts back when the user is finished
        function TaskListViewModel() {
            var self = this;

            // Properties
            self.tasks = ko.observableArray([]);
            self.newTaskText = ko.observable();
            self.incompleteTasks = ko.computed(function () {
                return ko.utils.arrayFilter(self.tasks(), function (task) { return !task.isDone() });
            });

            // Methods
            self.addTask = function () {
                self.tasks.push(new Task({ Title: this.newTaskText() })); // Note that this is case sensitive. The case of the title must match the case in the Task definition - fussy isn't it...
                self.newTaskText("");
            };
            self.removeTask = function (task) { self.tasks.remove(task) };

            // Send final state to server to be preserved
            // Note that we use an Ajax post here is that appears to integrate best with MVC
            self.save = function () {
                $.ajax("/TaskList/Save/", {
                    type: "POST",
                    contentType: "application/json",
                    data: ko.toJSON({ tasks: self.tasks })
                })
                .done(function (data) {
                    // Just to see if anything is returned
                    alert(data);
                })
                .fail(function () {
                    alert('Data not saved.');
                });
            };

            // Load initial state from server, convert it to Task instances, then populate self.tasks
            // Again, note that we use an Ajax post here is that appears to integrate best with MVC
            $.ajax({
                url: "/TaskList/GetTasks/",
                type: "POST",
                data: ko.toJSON({ tasks: self.tasks })
            })
            .done(function (data) {
                // Convert data to array of tasks using JQuery map function. See http://api.jquery.com/jQuery.map/
                var mappedTasks = $.map(data, function (item) { return new Task(item) });
                self.tasks(mappedTasks);
            })
            .fail(function () {
                alert('Initial data not received.');
            });

        }

        ko.applyBindings(new TaskListViewModel());

    </script>

End Section