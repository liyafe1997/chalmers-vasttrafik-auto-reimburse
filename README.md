# chalmers-vasttrafik-auto-reimburse

Send Västtrafik receipt to university email address automatically for reimbursement.

Based on https://chalmersstudentkar.se/8446-2/

# Usage

## 1. Enable Outlook run script by rules function
For security reason, Outlook disabed this feature in default.

Basically you can modify registety to enable this function. But different outlook version in different path.

So you can just Google "Outlook 2016(or your version) enable rule run script" to get how to do that.

## 2. Enable Developer Menu
Go to FILE->Options->Customize Ribbon, Then check "Developer"

![1](https://user-images.githubusercontent.com/18359157/143296666-ef52bd4e-23b3-4d36-b589-d713329f2764.jpg)

## 3. Import vba script
Go to DEVELOPER memu -> Visual Basic. Right-click "Project1" -> Insert -> Module. Then copy the content in the Module1.bas to the Module1.

(Or, you can just click "Import File" and then import Module1.bas in this repo).

![2](https://user-images.githubusercontent.com/18359157/143303235-b7ee79d8-c005-49c7-931a-18ae2848b1a4.jpg)

## 4. Fill in your information to the variable

MyEmailAddress is your Chalmers email address.

TicketType is the string appear in the receipt. It use for only forward the specific type of ticket. Please double check it while you use it. I am not sure whether it will change in the future.

ＭyName is your full name. This will appear on the email be forwarded.

MyStudentCardNumber is the card number, it will appear on the email be forwarded.

TargetMailAddress is the address forward to. You can use your own email to test firstly, then change to university's email.

![3](https://user-images.githubusercontent.com/18359157/143303303-bb7c3a02-721b-4f75-8274-4bc23f5016c4.jpg)

## 5. Setup rule
Click FILE->Rules and Alerts

![4](https://user-images.githubusercontent.com/18359157/143304489-b64cbfc2-a4d7-4ee1-9480-3ef3058ea72f.jpg)

Click "New Rule..." -> Apply rule on messages I receive. Then Next.

![7](https://user-images.githubusercontent.com/18359157/143304717-a16e7679-f4b8-480e-a02f-8fd572c07e9b.jpg)

Check "with specific works in the subject", set the specific words to "Betalningsbekräftelse - Västtrafik To Go"

(Same please double check the subject name, I not sure it has changed or not when you read this article in the future.)

![5](https://user-images.githubusercontent.com/18359157/143304774-18786139-32d0-4c6d-9aab-e885a17164de.jpg)

Next, check "run a script", and select the script "ForwardVasttrafik".

![6](https://user-images.githubusercontent.com/18359157/143305070-c2506b66-8049-4e26-a0fe-5f11e9320f48.jpg)

Done.

## One more little Tips
You can enable "Send Receipts" in Västtrafik app settings. The receipt will be sent to your email automatically when you buy a ticket. Then the script will forward the receipt to the school's email automatically.

![8](https://user-images.githubusercontent.com/18359157/143305474-df32d8cb-1595-4862-89d6-10085fb3d381.jpg)
