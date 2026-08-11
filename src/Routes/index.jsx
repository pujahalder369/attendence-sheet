import { useRoutes } from 'react-router-dom';
import PublicRoute from './PublicRoute';

const MainRoutes = () => {
    const routes = [...PublicRoute];
    const element = useRoutes(routes);
    return element;
}

export default MainRoutes;