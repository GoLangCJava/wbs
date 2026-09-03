# One-Abbott Shared Application Platform

## 1. 项目背景与目标

### 1.1 当前挑战
- 重复建设
- 数据孤岛
- 用户体验割裂
- 多套运营后台
- AI能力无法共享

### 1.2 建设目标
- One-Abbott统一平台
- 能力复用
- 数据复用
- AI复用
- 降本增效

---

## 2. 总体架构原则

### Capability First
### Data First
### AI First
### Micro Frontend
### Event Driven
### Plug-in Architecture

---

## 3. 企业架构（Enterprise Architecture）

### 3.1 架构分层

Experience Layer
↓
Application Layer
↓
Capability Layer
↓
Data Layer
↓
AI Layer
↓
Cloud Layer

### 3.2 目标架构图

（此处放 Gartner EA 架构图）

---

## 4. 业务架构（Business Architecture）

### 4.1 HCP全生命周期

产品认知
↓
学术学习
↓
会议互动
↓
销售拜访
↓
售后服务
↓
持续运营

### 4.2 八大业务域

① 产品介绍

② 学术内容

③ 会议管理

④ 互动积分

⑤ 拜访管理

⑥ 客户管理

⑦ 销售绩效

⑧ 报修服务

---

## 5. 业务能力映射模型

Business Capability
↓
Application Capability
↓
Technology Capability

---

### 5.1 产品介绍

#### Business Capability
产品传播管理

#### Application Capability
Product Content Capability

#### Technology Capability
CMS
DAM
Search
Publishing

---

### 5.2 学术内容

#### Business Capability
医学教育管理

#### Application Capability
Academic Content Capability

#### Technology Capability
Knowledge Hub
LMS
AI Search
Recommendation Engine

---

### 5.3 会议管理

#### Business Capability
学术活动管理

#### Application Capability
Event Capability

#### Technology Capability
Workflow Engine
Streaming Platform
Calendar Service
Notification Service

---

### 5.4 互动积分

#### Business Capability
用户运营

#### Application Capability
Engagement Capability

#### Technology Capability
Rules Engine
Marketing Automation
Loyalty Platform

---

### 5.5 拜访管理

#### Business Capability
销售执行管理

#### Application Capability
Sales Execution Capability

#### Technology Capability
CRM
Mobile Platform
GPS Check-in
Offline Sync

---

### 5.6 客户管理

#### Business Capability
客户资产管理

#### Application Capability
Customer Capability

#### Technology Capability
MDM
Identity Platform
Customer360

---

### 5.7 销售绩效

#### Business Capability
销售经营管理

#### Application Capability
Sales Performance Capability

#### Technology Capability
BI
Analytics
Forecast Engine
KPI Engine

---

### 5.8 报修服务

#### Business Capability
售后服务管理

#### Application Capability
Service Capability

#### Technology Capability
Field Service
Ticket Engine
IoT Platform
Knowledge Base

---

## 6. 应用架构（Application Architecture）

### 6.1 当前模式

ADC
EPD
MD
APOC

各自建设内容、会议、积分

---

### 6.2 目标模式

Capability Platform

Content
Event
Customer
Sales
Service
AI

Build Once
Reuse Everywhere

---

### 6.3 微前端架构

Portal Shell

Product MF

Content MF

Event MF

Sales MF

Service MF

Customer MF

AI MF

Plugin MF

---

### 6.4 微前端技术栈

React

Module Federation

Single SPA

Next.js

---

## 7. 能力中心设计（Capability Hub）

### Content Hub

产品中心

学术中心

文献中心

课程中心

---

### Event Hub

会议管理

直播

签到

积分

---

### Customer Hub

HCP

HCO

Customer360

标签引擎

---

### Sales Hub

拜访

协访

销量

绩效

---

### Service Hub

工单

SLA

备件

IoT

---

### AI Hub

AI问答

AI推荐

AI审核

AI助手

---

### Plugin Hub

术后随访

医学咨询

竞品情报

成本计算

差旅管理

---

## 8. 技术架构（Technology Architecture）

### Experience Layer

微信

小程序

企业微信

App

Portal

---

### API Layer

Azure API Management

---

### Service Layer

Content Service

Event Service

Customer Service

Sales Service

Service Service

AI Service

---

### Event Layer

Azure Event Hub

Azure Service Bus

---

### Platform Engineering

AKS

Container Registry

Terraform

CI/CD

Observability

---

## 9. 数据架构（HCP Pool）

### HCP Pool

OneID

HCP Master

HCO Master

Consent

Customer360

Tag Engine

Behavior Event

---

### 数据流

Channel
↓
Behavior Event
↓
HCP Pool
↓
Customer360
↓
AI Platform

---

## 10. AI架构

### AI Platform

Azure OpenAI

AI Search

RAG

Copilot

---

### AI能力

AI问答

AI推荐

AI客服

AI合规审核

AI运营分析

---

## 11. Azure Landing Zone映射

### Identity

Microsoft Entra ID

External ID

---

### Compute

AKS

Container Apps

---

### API

Azure API Management

---

### Messaging

Event Hub

Service Bus

---

### Storage

Azure SQL

Cosmos DB

Blob Storage

---

### Data

Data Lake

Fabric

Databricks

---

### AI

Azure OpenAI

AI Search

Copilot

---

### Monitoring

Azure Monitor

App Insights

Log Analytics

---

## 12. 三阶段实施路线

### Phase 1

统一底座

OneID

CMS

Event Platform

HCP Pool

---

### Phase 2

能力复用

Micro Frontend

Capability Hub

Plugin Framework

Customer360

---

### Phase 3

AI赋能

AI Search

AI Recommendation

AI Copilot

AI Compliance

---

## 13. 最终目标架构

Channels

↓

Micro Frontend

↓

Capability Hub

↓

HCP Pool

↓

AI Platform

↓

Azure Landing Zone

### 核心价值

Capability Reuse

Data Reuse

AI Reuse

Build Once
Reuse Everywhere

