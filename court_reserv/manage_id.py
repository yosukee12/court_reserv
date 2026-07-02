# -*- coding: utf-8 -*-
try:
    from .config import load_config
    from .services import IdManagerService
except Exception:
    from config import load_config
    from services import IdManagerService

config = load_config()
id_manager_service = IdManagerService(config)

class Manage_Id():
    @staticmethod
    def get_id_dict_from_csv(csv_file_path):
        return id_manager_service.load_accounts(csv_file_path)

    @staticmethod
    def output_csv_from_id_dict(id_dict, output_file_path):
        return id_manager_service.save_accounts(id_dict, output_file_path)

    @staticmethod
    def get_alive_dead_id_dict(id_dict):
        return id_manager_service.check_account_validity(id_dict)
