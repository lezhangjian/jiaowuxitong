import requests
from lxml import etree
import xlwt
import re
from PIL import Image
import pytesser3


class Jiaowuchu():
    def __init__(self):
        # 保证之后的request请求使用的相同的cookies
        self.s = requests.session()
        self.Cookie = self.s.get('http://jwxt.upc.edu.cn/jwxt/').cookies
        self.Cookie = 'JSESSIONID' + '=' + self.Cookie['JSESSIONID']
        self.headers = {
            'Cookie': self.Cookie,
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)'
        }
        self.s.get('http://jwxt.upc.edu.cn/jwxt/', headers=self.headers)
        self.username = '1606050220'
        self.password = '*******'
        self.user = '无'

    # 登陆教务系统,并下载课表数据
    def login(self):
        # 下载验证码
        with open('captcha.jpg', 'wb') as f:
            f.write(self.s.get('http://jwxt.upc.edu.cn/jwxt/verifycode.servlet').content)
        # 手动输入验证码
        # captcha_code = input('输入验证码>>')
        # 自动识别验证码
        captcha_code = self.recognize_captcha().strip()
        data = {
            'USERNAME': self.username,
            'PASSWORD': self.password,
            'RANDOMCODE': captcha_code,
        }
        url = 'http://jwxt.upc.edu.cn/jwxt/Logon.do?method=logon'
        # 提交账号密码,,验证码,,并登陆
        req =  self.s.post(url, data=data)
        page = etree.HTML(req.text)
        error = page.xpath('//span[@id="errorinfo"]/text()')
        if error:
            print(error)
            self.login()
    
        url = 'http://jwxt.upc.edu.cn/jwxt/framework/main.jsp'
        req = self.s.get(url)
        page = etree.HTML(req.text)
        user = page.xpath('//title/text()')
        self.user = re.search(r'(.*?)\[', user[0]).groups(1)
        url = 'http://jwxt.upc.edu.cn/jwxt/Logon.do?method=logonBySSO'
        self.s.post(url)
    
    def course_table(self):
        excel = xlwt.Workbook(encoding='utf-8')
        week_list = ['0','一','二','三','四','五','六','七','八','九','十','十一','十二','十三','十四','十五','十六','十七']
        for k in range(1,18):
            url = 'http://jwxt.upc.edu.cn/jwxt/tkglAction.do?method=goListKbByXs&sql=&xnxqh=2018-2019-2&zc={}&xs0101id={}'
            req = self.s.get(url.format(k, self.username))
            page = etree.HTML(req.text)
    
            sh = '{}周'.format(week_list[k])
            sheet = excel.add_sheet(sh)
            style = xlwt.XFStyle()  # 初始化样式
            font = xlwt.Font()  # 创建字体
            style.alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT  # 自动换行
            style.font = font  # 设定样式
    
            for r in range(1, 7):
                font1 = xlwt.Font()  # Create Font
                font1.height = 20 * 20  # 字体大小
                style1 = xlwt.XFStyle()  # Create Style
                style1.alignment.vert = 0x01  # 字体居中
                style1.font = font1
                title = str(r * 2 - 1) + '~' + str(r * 2)
                sheet.write(r, 0, title, style1)
    
            date_dict = {
                '1': '星期一',
                '2': '星期二',
                '3': '星期三',
                '4': '星期四',
                '5': '星期五',
                '6': '星期六',
                '7': '星期日',
            }
            for c in range(1, 8):
                sheet.row(0).height_mismatch = True
                sheet.row(0).height = 20 * 30
                font2 = xlwt.Font()  # Create Font
                font2.height = 20 * 14  # 字体大小
                font2.bold = True  # 粗体
                style2 = xlwt.XFStyle()  # Create Style
                style2.font = font2
                sheet.write(0, c, date_dict[str(c)], style2)
    
            for r in range(1, 7):
                sheet.row(r).height_mismatch = True
                sheet.row(r).height = 20 * 120
                for c in range(1, 8):
                    sheet.col(c).width = 256 * 13
                    id = str(r) + '-' + str(c) + '-' + '2'
                    path = '//div[@id="{}"]//text()'
                    data = page.xpath(path.format(id))
                    sheet.write(r, c, data, style)
            name = self.user[0]+'的课表.xls'
            excel.save(name)
    
    def recognize_captcha(self):
        print('正在识别验证码!!!')
        # 灰度化
        img = Image.open('captcha.jpg')
        img = img.convert('L')
        # 二值化
        pixdata = img.load()
        width, height = img.size
        threshold = sum(img.getdata()) / (width * height)  # 计算图片的平均阈值
        # 遍历所有像素，大于阈值的为白色
        for y in range(height):
            for x in range(width):
                if pixdata[x, y] < threshold:
                    pixdata[x, y] = 0
                else:
                    pixdata[x, y] = 255
        # 去掉黑边
        for y in range(height):
            pixdata[0, y] = 255
            pixdata[width - 1, y] = 255
        for x in range(width):
            pixdata[x, 0] = 255
            pixdata[x, height - 1] = 255
        # 降噪
        N = 2
        for y in range(1, height - 1):
            for x in range(1, width - 1):
                count = 0
                if pixdata[x, y - 1] == 255:  # 上
                    count = count + 1
                if pixdata[x, y + 1] == 255:  # 下
                    count = count + 1
                if pixdata[x - 1, y] == 255:  # 左
                    count = count + 1
                if pixdata[x + 1, y] == 255:  # 右
                    count = count + 1
                if count > N:
                    pixdata[x, y] = 255  # 设置为白色
        captcha_code = pytesser3.image_to_string(img).strip()
        print(captcha_code)
        return captcha_code


if __name__ == '__main__':
    run = Jiaowuchu()
    run.login()
    print('登陆成功!!!')
    print('正在下载课表...')
    run.course_table()
    print('下载完成!!!')
    # print(run.user)
