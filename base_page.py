# coding  : utf8
# @Date   : 2020/7/31 - 16:39
# @Author : Ding Ning

from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException


class BasePage:
    def __init__(self):
        # 指定Chromedriver路径
        option = webdriver.ChromeOptions()
        option.add_argument("--user-data-dir=" + r"C:/Users/Administrator/AppData/Local/Google/Chrome/User Data/")
        self.driver = webdriver.Chrome(chrome_options=option)
        # self.driver.maximize_window()
        self.driver.implicitly_wait(20)

    # 定义一个私有类方法，其他类不可调用
    def open(self, url):
        self.driver.get(url)
        self.driver.maximize_window()

    # 重写find_element方法，参数为可变数量参数，保证元素可见
    def find_element(self, by, locator, timeout=20):
        """
        定位单个元素
        :param by: 定位方式 eg:By.ID
        :param locator: 定位表达式
        :param timeout: 显示等待超时时间
        :return:
        """
        try:
            element = WebDriverWait(self.driver, timeout). \
                until(lambda driver: driver.find_element(by, locator))
        except (NoSuchElementException, TimeoutException) as e:
            raise e
        else:
            return element

    def find_elements(self, by, locator, timeout=20):
        """
        定位一组元素
        :param by: 定位方式 eg:By.ID
        :param locator: 定位表达式
        :param timeout: 显示等待超时时间
        :return:
        """
        try:
            elements = WebDriverWait(self.driver, timeout). \
                until(lambda driver: driver.find_elements(by, locator))
        except (NoSuchElementException, TimeoutException) as e:
            raise e
        else:
            return elements

    # 定义script方法，用于执行js脚本
    def script(self, src):
        self.driver.execute_script(src)

    # 定义页面跳转方法
    def switch_frame(self, loc):
        return self.driver.switch_to.frame(loc)

    # 重新定义send_key方法，为了保证搜索按钮存在，输入框中的值要清空
    def send_keys(self, loc, value, clear_first=True, click_first=True):
        # noinspection PyBroadException
        try:
            # 用getattr方法实现self.loc
            loc = getattr(self, f'_{loc}')
            # 是否存在搜索按钮
            if click_first:
                self.find_element(*loc).click()
            # 清空搜索框中的值，输入值
            if clear_first:
                self.find_element(*loc).clear()
                self.find_element(*loc).send_keys(value)
        except Exception:
            print(f'页面上未找到{loc}元素')
