from docxtpl import DocxTemplate
from tkinter import *
from tkinter import messagebox


class Suggestions:
    def __init__(self):
        self.root = Tk()
        self.frame = Frame(self.root)
        self.frame.pack()
        self.root.geometry("680x480")


        # add input box for user to enter text
        name_var = StringVar()
        L1 = Label(self.frame, text = "Client Name")
        L1.pack( side = LEFT)
        E1 = Entry(self.frame, textvariable=name_var, bd=5)
        E1.pack(side = RIGHT)

        # top frame will contain users name
        topFrame = Frame(self.root)
        topFrame.pack(side=TOP)

        # middle frame will contact check boxes
        middleFrame = Frame(self.root)
        middleFrame.pack(side=TOP)

        # bottom frame will contain submit button
        bottomFrame = Frame(self.root)
        bottomFrame.pack(side=TOP)

        # variable that will capture users input from check boxes
        Live_Chat = IntVar()
        Contact_Form = IntVar()
        Responsiveness = IntVar()
        Product_Search = IntVar()

        C1 = Checkbutton(middleFrame, text="Live Chat", variable=Live_Chat, onvalue=1, offvalue=0, height=2, width=20)
        C2 = Checkbutton(middleFrame, text="Contact Form", variable=Contact_Form, onvalue=1, offvalue=0, height=2, width=20)
        C3 = Checkbutton(middleFrame, text="Responsiveness", variable=Responsiveness, onvalue=1, offvalue = 0, height=2, width=20)
        C4 = Checkbutton(middleFrame, text="Product Search", variable=Product_Search, onvalue=1, offvalue = 0, height=2, width=20)

        C1.pack()
        C2.pack()
        C3.pack()
        C4.pack()


        selected_list = [Live_Chat, Contact_Form, Responsiveness, Product_Search]


        def suggested_improvements():
            name = E1.get()
            # get the values from the checkbox(es) selected by the user
            lc = Live_Chat.get()
            cf = Contact_Form.get()
            rwd = Responsiveness.get()
            ps = Product_Search.get()

            live_chat = "Live Chat" + "\n\n" \
                                      "Live chat software enables you to have real-time conversations with your customers while they are on your " \
                                      "website. The American Marketing Association found that B2B companies who used live chat see, on " \
                                      "average, a 20% increase in conversions!" + "\n" + "(https://www.ama.org/Documents/how-b2b-marketers-leveraging-live-chat-increase-sales.pdf)" + \
                        "\n" + "eMarketer report found that 35% more people converted from visitor to customer after using a website " \
                               "live chat." + "\n" + "(https://www.emarketer.com/Article/How-Helpful-Live-Chat/1007235)"

            contact_form = "Contact Form" + "\n\n" \
                                            "A contact form allows you to gather customer data and grow a subscriber list while allowing visitors to " \
                                            "communicate with you during non-business hours.  A contact form ensure you get the information you want, " \
                                            "then develop marketing and follow-up programs. "

            responsive_web_design = "Responsive Web Design" + "\n\n" \
                                                              "Responsive web design leads to higher conversion rates.  According to Google, a business is " \
                                                              "likely to lose 61% percent of visitors if its mobile website is difficult to navigate. " \
                                                              "However, if the users enjoy a positive interaction with a website, they are 67% percent more " \
                                                              "likely to convert." \
                                    + "\n" + "https://www.business.com/articles/how-responsive-web-design-helps-get-you-more-conversions/"

            product_search = "Product Search" + "\n\n" \
                                                "If you’ve ever visited www.amazon.com, you know how useful their product search functions are. " \
                                                "Implementing this type of functionality is easy for any website and doing so creates a " \
                                                "user-friendly experience. Given the product range you offer, a product search function allows " \
                                                "visitors the ability to intuitively and quickly find the information about the products you offer."

            # content that will be sent to .docx file
            content = {'chat': '', 'form': '', 'mobile': '', 'search': '',
                       'live_chat': '', 'contact_form': '', 'responsive_web_design': '', 'product_search_function': '',
                       'name': name}

            # array of strings containing facts about each website improvement
            suggestions_list = [live_chat, contact_form, responsive_web_design, product_search]

            # live chat
            if lc > 0:
                content['chat'] = 'Live Chat,'
                content['live_chat'] = live_chat

            # contact form
            if cf > 0:
                content['form'] = 'Contact Form,'
                content['contact_form'] = contact_form

            # responsive web design
            if rwd > 0:
                content['mobile'] = 'Responsive Web Design,'
                content['responsive_web_design'] = responsive_web_design

            # search product function
            if ps > 0:
                content['search'] = 'Product Search Function,'
                content['product_search_function'] = product_search

            # add content to docx template
            doc = DocxTemplate('C:/Users/gtohi/OneDrive/test.docx')
            doc.render(content)
            doc.save("C:/Users/gtohi/OneDrive/generated.docx")

            # show alert message to screen

            msg = messagebox.showinfo("Suggestions", "Success")





        B = Button(bottomFrame, text="Submit", command=suggested_improvements)
        B.pack()

        # Code to add widgets will go here...
        self.root.mainloop()


app = Suggestions()



