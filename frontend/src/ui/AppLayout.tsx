import { BankOutlined } from '@ant-design/icons';
import { Layout, Menu, Typography } from 'antd';
import { Outlet, useNavigate } from 'react-router-dom';

const { Header, Sider, Content } = Layout;

export function AppLayout() {
  const navigate = useNavigate();

  return (
    <Layout className="app-shell">
      <Sider width={248} className="app-sidebar">
        <Typography.Title level={4} className="app-logo">
          내부 관리
        </Typography.Title>
        <Menu
          mode="inline"
          selectedKeys={['billing']}
          onClick={() => navigate('/')}
          items={[{ key: 'billing', icon: <BankOutlined />, label: '청구자료 생성' }]}
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
