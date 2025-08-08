import threading

class SharedData:
    """A thread-safe class to hold shared application data."""

    def __init__(self):
        self.lock = threading.Lock()

        # Parcel Info
        self.parcel_pin = ""
        self.parcel_address = ""
        self.parcel_owner = ""
        self.parcel_city_state_zip = ""
        self.parcel_legal_description = ""

        # Tax Info
        self.tax_2024_total = ""
        self.tax_2024_paid_status = ""
        self.tax_2025_estimated = ""

        # Title Chain Info
        self.title_chain_results = []

        # Manual Fields
        self.lender = ""
        self.borrower = ""
        self.loan_amount = ""
        self.writer = ""
        self.date = ""
        self.notes = ""

    def get_data(self, key):
        with self.lock:
            return getattr(self, key, None)

    def set_data(self, key, value):
        with self.lock:
            setattr(self, key, value)

    def update_data(self, data_dict):
        with self.lock:
            for key, value in data_dict.items():
                if hasattr(self, key):
                    setattr(self, key, value)
