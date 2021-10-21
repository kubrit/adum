### ADUM - Active Directory User Management

Application maded specialy for HR Department to set basic user/account information.

Search for user and modify attributes. There is a feature as `custom-attr` that you cann setup. If you dont have any just comment that out.

> Note: remember to set permissions first on active directory group for only those attributes that are used in script.
```sh
    sAMAccountName,
    displayName,
    mail,
    StreetAddress,
    postOfficeBox,
    physicalDeliveryOfficeName,
    l,
    st,
    PostalCode,
    c,
    pager,
    custom-attr,
    mobile,
    title,
    department,
    company```
