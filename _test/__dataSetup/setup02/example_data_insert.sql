--delete from entities5dim_tab;
insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('HOLD', 'Supplier', 'Supp Dude', 'Brakes', 'EUR', 1012.69, 202.54);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('LTD', 'Supplier', 'George Doors', 'Gears', 'GBP', 1200, 0);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('HOLD', 'Customer', 'Car Bits', 'Brakes', 'GBP', 600.78, 120.16);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('LTD', 'Supplier', 'Green Engines', 'Brakes', 'GBP', 354, 0);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('LTD', 'Customer', 'Ford', 'Interior', 'SEK', 789.54, 157.9);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('LTD', 'Customer', 'Ford', 'Brakes', 'GBP', 416.8, 83.36);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('HOLD', 'Customer', 'Volvo', 'Interior', 'GBP', 168.56, 0);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('LTD', 'Customer', 'Ford', 'Gears', 'USD', 56.4, 0);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('LLB', 'Customer', 'Car Bits', 'Gears', 'EUR', 333.34, 66.67);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('LLB', 'Supplier', 'George Doors', 'Interior', 'EUR', 56.4, 11.28);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('LLB', 'Supplier', 'George Doors', 'Interior', 'EUR', 30.4, 6.08);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('LLB', 'Customer', 'Ford', 'Brakes', 'USD', 700.3, 0);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('HOLD', 'Supplier', 'Supp Dude', 'Gears', 'SEK', 450, 90);

insert into entities5dim_tab (company, identity_type, identity, category, currency, amount, tax)
values ('LLB', 'Customer', 'Ford', 'Interior', 'SEK', 650, 0);

commit;
