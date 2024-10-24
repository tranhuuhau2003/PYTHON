
class StudentController:
    def __init__(self, model, view):
        self.model = model
        self.view = view

    def load_data(self, filepath):
        try:
            self.model.load_from_excel(filepath)
            students = self.model.fetch_all_students()
            self.view.populate_table(students)
        except Exception as e:
            self.view.show_error(str(e))

    def on_close(self):
        self.model.close()
