import requests
from library.file_operation import FileOperation


class DbotOPeration:
    def __init__(self):
        self.fo = FileOperation()
        self.log = self.fo.log

    def create_dbot_sniper(self, user_snip_config, contract):
        """
        dbot狙击
        :return:
        """
        xkey_id = user_snip_config['xkey']
        buy_amount = user_snip_config['buy_amount']
        jito_tip = user_snip_config['jitoTip']
        max_slippage = user_snip_config['maxSlippage']
        url = 'https://api-bot-v1.dbotx.com/automation/snipe_order'
        header = {
            'X-API-KEY': xkey_id
        }
        params = {
            "enabled": True,
            "chain": "solana",
            "token": contract,
            "walletId": "main01",
            "expireDelta": 3600000,
            "buySettings": {
                "buyAmountUI": buy_amount,
                "priorityFee": "0.01",
                "gasFeeDelta": 5,
                "maxFeePerGas": 100,
                "jitoEnabled": True,
                "jitoTip": jito_tip,
                "maxSlippage": max_slippage,
                "minLiquidity": 5000,
                "concurrentNodes": 3,
                "retries": 1
            }
        }
        res = requests.get(url=url, headers=header, params=params)
        if res.status_code == 200:
            rep_json = res.json()
            self.log.info(f'狙击成功：{rep_json}')
        else:
            self.log.info(f'狙击失败：{res.text}')

    def create_sell_order(self, user_snip_config, contract):
        """
        创建卖单
        :param user_snip_config:
        :param contract:
        :return:
        """
        xkey_id = user_snip_config['xkey']
        buy_amount = user_snip_config['buy_amount']
        jito_tip = user_snip_config['jitoTip']
        max_slippage = user_snip_config['maxSlippage']
        url = 'https://api-bot-v1.dbotx.com/automation/limit_orders'
        header = {
            'X-API-KEY': xkey_id
        }
        params = {
            "enabled": True,
            "chain": "solana",
            "token": contract,
            "walletId": "main01",
            "expireDelta": 3600000,
            "buySettings": {
                "buyAmountUI": buy_amount,
                "priorityFee": "0.01",
                "gasFeeDelta": 5,
                "maxFeePerGas": 100,
                "jitoEnabled": True,
                "jitoTip": jito_tip,
                "maxSlippage": max_slippage,
                "minLiquidity": 5000,
                "concurrentNodes": 3,
                "retries": 1
            }
        }
        res = requests.get(url=url, headers=header, params=params)
        if res.status_code == 200:
            rep_json = res.json()
            self.log.info(f'狙击成功：{rep_json}')
        else:
            self.log.info(f'狙击失败：{res.text}')

