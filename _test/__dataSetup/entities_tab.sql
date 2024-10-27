-- Create table
create table ENTITIES_TAB
(
  identity_type VARCHAR2(50) not null,
  identity      VARCHAR2(20) not null,
  currency      VARCHAR2(3),
  amount        NUMBER
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
alter table ENTITIES_TAB
  add constraint ENTITIES_PK primary key (IDENTITY_TYPE, IDENTITY)
  using index 
  tablespace USERS
  pctfree 10
  initrans 2
  maxtrans 255;
