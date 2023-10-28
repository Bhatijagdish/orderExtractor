from typing import List

from lib.utils import trim_channel_name
from lib.xl_csv_reader import ReadFile


class Orders(ReadFile):

    def __init__(self, filepath: str):
        super().__init__(filepath)
        self.get_sheet_using_sheet_name(0)

    def list_sku_code(self) -> List:
        return self.get_column_data_by_header("Listing Sku Code")

    def get_channel_names(self) -> List:
        return self.get_column_data_by_header("Channel Name")

    def total_sales(self) -> List:

        return [float(item) for item in self.get_column_data_by_header("Total")]

    def order_statuses(self) -> List:
        return self.get_column_data_by_header("Order Status")

    def shipping_fee(self) -> List:
        return [abs(float(item)) for item in self.get_column_data_by_header("Shipping Fee")]

    def service_tax(self) -> List:
        return [abs(float(item)) for item in self.get_column_data_by_header("Service Tax")]

    def channel_fees(self) -> List:
        return [abs(float(item)) for item in self.get_column_data_by_header("Total Channel Fees")]

    def settlement_amounts(self) -> List:
        return [float(item) for item in self.get_column_data_by_header("Settlement Amount")]

    def tax_rates(self) -> List:
        return [float(item) for item in self.get_column_data_by_header("Tax Rate")]

    # def return_dates(self) -> List:
    #     return [abs(item) for item in self.get_column_data_by_header("Return Date")]

    def quantities(self) -> List:
        return [float(item) for item in self.get_column_data_by_header("Qty")]

    def get_month(self) -> str:
        date: str = self.input_sheet["A2"].value
        return date[:3].upper()

    def trim_channel_names(self):
        return [trim_channel_name(item) for item in set(self.get_channel_names())]

