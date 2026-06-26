import { BankOutlined, LogoutOutlined } from '@ant-design/icons';
import { Button, Layout, Menu, Typography } from 'antd';
import { Outlet, useNavigate } from 'react-router-dom';

import { useAuthStore } from '../stores/authStore';

const { Header, Sider, Content } = Layout;

export function AppLayout() {
  const navigate = useNavigate();
  const logout = useAuthStore((state) => state.logout);

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
          <Button
            icon={<LogoutOutlined />}
            onClick={() => {
              logout();
              navigate('/login');
            }}
          >
            로그아웃
          </Button>
        </Header>
        <Content className="app-content">
          <Outlet />
        </Content>
      </Layout>
    </Layout>
  );
}
