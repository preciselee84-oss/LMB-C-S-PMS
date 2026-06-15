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
          위탁사업장 관리
        </Typography.Title>
        <Menu
          mode="inline"
          selectedKeys={['dashboard']}
          onClick={() => navigate('/')}
          items={[{ key: 'dashboard', icon: <BankOutlined />, label: '관리 대시보드' }]}
        />
      </Sider>
      <Layout>
        <Header className="app-header">
          <Typography.Text strong>전도금 요청-품의-이체 자동화 MVP</Typography.Text>
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
