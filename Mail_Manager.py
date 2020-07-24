from O365 import Account
import os

credentials = ('74424fcf-55d7-4e15-99d7-1663c0ba2e94',)
account = Account(credentials, auth_flow_type='public')


class mailbox_actions:
    """用来对邮箱进行操作"""

    def __init__(self, choice=None):

        # 这是我的应用ID和机密，但是公共客户端流还是没有实现，文档里说下面改为
        # self.account = Account(self.credentials, auth_flow_type=‘pubic')
        # （上面是我的理解）就行了但他报错
        self.account = account
        self.credentials = credentials
        self.choice = choice
        self.scopes = ['basic', 'message_all']  # 请求权限

    def check_if_authenticated(self):
        """检查是否有用户登录，若无，则请求登录"""
        if not self.account.is_authenticated:  # 检查是否登录
            # 请求登录
            self.account.authenticate(scopes=self.scopes)

    def Read_email(self):
        """遍历邮件
        limit 表示加载多少个，微软官方一次API调用只返回999个，
        而O365模块默认25个，只有limit>25时utils分页功能才生效
        batch批处理表示加载多少次，就是往后加载
        limit=2000, batch=10 = limit=2000
        但是分为10次加载。"""

        mailbox = self.account.mailbox()

        inbox = mailbox.inbox_folder()
        for messages in inbox.get_messages(limit=200, batch=100):  # 下面的都是utils分页的
            print(messages)
        for messages in mailbox.junk_folder().get_messages(limit=200, batch=100):
            print(messages)
        for messages in mailbox.deleted_folder().get_messages(limit=200, batch=100):
            print(messages)
        for messages in mailbox.drafts_folder().get_messages(limit=2000):
            print(messages)
        for messages in mailbox.sent_folder().get_messages(limit=2000, batch=10):
            print(messages)
        os.system("pause")

    # 准备加入选择进入哪个文件夹

    def start(self):
        """应用入口"""
        self.choice = input('进入邮箱还是日历？(E/C)')

        if self.choice == 'E':
            self.choice = input('看邮件还是写邮件？(R/W)')
            if self.choice == 'R':
                mailbox_actions().Read_email()
            elif self.choice == 'W':
                print('开发中，请稍后')
            else:
                mailbox_actions().start()
        elif self.choice == 'C':
            print('开发中，敬请期待')
        else:
            mailbox_actions().start()


mailbox_actions().check_if_authenticated()
mailbox_actions().start()
