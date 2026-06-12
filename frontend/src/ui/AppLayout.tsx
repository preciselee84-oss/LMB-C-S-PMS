import { BarChartOutlined, DatabaseOutlined, LogoutOutlined } from '@ant-design/icons';
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
          LMB 실적관리
        </Typography.Title>
        <Menu
          mode="inline"
          defaultSelectedKeys={['dashboard']}
          items={[
            { key: 'dashboard', icon: <BarChartOutlined />, label: '대시보드' },
            { key: 'data', icon: <DatabaseOutlined />, label: '데이터 관리' },
          ]}
        />
      </Sider>
      <Layout>
        <Header className="app-header">
          <Typography.Text strong>정식 전환 프로젝트</Typography.Text>
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

