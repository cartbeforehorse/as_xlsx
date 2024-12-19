-- Create table
create table ENTITIES5DIM_TAB
(
  company       VARCHAR2(20) not null,
  identity_type VARCHAR2(50) not null,
  identity      VARCHAR2(20) not null,
  category      VARCHAR2(30) not null,
  currency      VARCHAR2(3)  not null,
  amount        NUMBER,
  tax           NUMBER
)
tablespace USERS
  pctfree 10
  initrans 1
  maxtrans 255
  storage
  (
    initial 64K
    minextents 1
    maxextents unlimited
  );
-- Create/Recreate primary, unique and foreign key constraints 
