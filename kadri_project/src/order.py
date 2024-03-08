class order:
    def __init__(self, dep: str, post: str, name: str,
                 status: str, new_department: str, data: str, number: str,
                 dt: str, sd: str):
        self.department = dep
        self.post = post
        self.name = name
        self.status = status
        self.new_department = new_department
        self.data = data
        self.order_number = number
        self.document_type = dt
        self.statement_data = sd

    def get_order_info_str(self):
        """
        Возвращает строку со всей информацией по приказу.

        Вроде не нужная функция
        :return: str
        """

        return f"{self.department}, {self.post}, {self.name}, {self.status}, " \
               f"{self.new_department}, {self.data}, {self.order_number}"

    def get_order_info_to_pandas(self):
        return [
            self.department,
            self.post,
            self.name,
            self.status,
            self.new_department,
            self.data,
            self.order_number,
            self.document_type,
            self.statement_data
        ]