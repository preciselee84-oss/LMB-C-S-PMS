import { BankOutlined, FormOutlined } from '@ant-design/icons';
import { Layout, Menu, Typography } from 'antd';
import { Outlet, useLocation, useNavigate } from 'react-router-dom';

const { Header, Sider, Content } = Layout;

export function AppLayout() {
  const navigate = useNavigate();
  const location = useLocation();
  const selectedKey = location.pathname.startsWith('/crm/visit-voc') ? 'visit-voc' : 'billing';

  return (
    <Layout className="app-shell">
      <Sider width={248} className="app-sidebar">
        <Typography.Title level={4} className="app-logo">
          내부 관리
        </Typography.Title>
        <Menu
          mode="inline"
          selectedKeys={[selectedKey]}
          onClick={({ key }) => navigate(key === 'visit-voc' ? '/crm/visit-voc' : '/')}
          items={[
            { key: 'billing', icon: <BankOutlined />, label: '청구자료 생성' },
            { key: 'visit-voc', icon: <FormOutlined />, label: 'CRM 방문 VOC' },
          ]}
        />
      </Sider>
      <Layout>
        <Header className="app-header">
          <Typography.Text strong>내부 업무 관리</Typography.Text>
        </Header>
        <Content className="app-content">
          <Outlet />
        </Content>
      </Layout>
    </Layout>
  );
}
