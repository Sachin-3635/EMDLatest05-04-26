import * as React from 'react';
import styles from './SonaEmd.module.scss';
import type { ISonaEmdProps } from './ISonaEmdProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { HashRouter as Router, Switch, Route, useLocation } from 'react-router-dom';

import 'bootstrap/dist/css/bootstrap.min.css';

import Sidebar from '../components/Pages/Sidebar';
import { InitiatorDashboard } from '../components/Pages/InitiatorDashboard';
import MyRequests from './Pages/MyRequests';
import ApprovalDashboard from './Pages/ApprovalDashboard';
import APTeamDashboard from './Pages/APTeamDashboard';
import ClosureApprovalARDashboard from './Pages/ClosureApprovalARDashboard';
import EMDClosureDashboard from './Pages/EMDClosureDashboard';
import TreasuryLandingPage from './Pages/TreasuryDashboard';
import EMDRequestForm from './Pages/EMDRequestForm';
import EMDRequestFormView from './Pages/EMDRequestFormView';
import MANACApprovalForm from './Pages/MANACApprovalForm';
import VouchingbyAPTeamForm from './Pages/VouchingbyAPTeamForm';
import UTRDetailEntry from './Pages/UTRDetailEntry';
import ClosureRequestForm from './Pages/ClosureRequestForm';
import ARClosureApprovalForm from './Pages/ARClosureApprovalForm';
import APTeamDashboardClosedByAR from './Pages/APTeamDashboardClosedByAR';
import EMDClosureRequestForm from './Pages/EMDClosureRequestForm';
import EMDRequestFormEdit from './Pages/EMDRequestFormEdit';
import ViewForm from './Pages/ViewForm';

const SonaEmd: React.FC<ISonaEmdProps> = (props) => {
  const {
    description,
    isDarkTheme,
    environmentMessage,
    hasTeamsContext,
    userDisplayName
  } = props;

  const location = useLocation(); // 

  const hideSidebar =
    location.pathname === "/NewRequest" ||
    location.pathname.startsWith("/ViewRequest/") ||
    location.pathname.startsWith("/EditRequest/") ||
    location.pathname.startsWith("/ApprovalRequest/");

  return (

    <div className="container-fluid" style={{ display: 'flex', width: '100%' }}>
      {!hideSidebar && <Sidebar {...props} />}
      <div className="main" style={{
        width: hideSidebar ? "100%" : "calc(100% - 250px)",
        transition: "width 0.3s ease"
      }}>
        <Switch>
          <Route exact path="/InitiatorDashboard" render={() => <InitiatorDashboard {...props} />} />
          <Route exact path="/" render={() => <MyRequests {...props} />} />
          <Route exact path="/ApprovalDashboard" render={() => <ApprovalDashboard {...props} />} />
          <Route exact path="/APTeamDashboard" render={() => <APTeamDashboard {...props} />} />
          <Route exact path="/APTeamDashboardClosedByAR" render={() => <APTeamDashboardClosedByAR {...props} />} />
          <Route exact path="/ClosureApprovalARDashboard" render={() => <ClosureApprovalARDashboard {...props} />} />
          <Route exact path="/EMDClosureDashboard" render={() => <EMDClosureDashboard {...props} />} />
          <Route exact path="/TreasuryLandingPage" render={() => <TreasuryLandingPage {...props} />} />
          <Route exact path="/EMDRequestForm" render={() => <EMDRequestForm {...props} />} />
          <Route exact path="/EMDRequestFormView" render={() => <EMDRequestFormView {...props} />} />
          <Route exact path="/EMDRequestFormEdit" render={() => <EMDRequestFormEdit {...props} />} />
          <Route exact path="/MANACApprovalForm" render={() => <MANACApprovalForm {...props} />} />
          <Route exact path="/VouchingbyAPTeamForm" render={() => <VouchingbyAPTeamForm {...props} />} />
          <Route exact path="/UTRDetailEntry" render={() => <UTRDetailEntry {...props} />} />
          <Route exact path="/ClosureRequestForm" render={() => <ClosureRequestForm {...props} />} />
          <Route exact path="/ARClosureApprovalForm" render={() => <ARClosureApprovalForm {...props} />} />
          <Route exact path="/EMDClosureRequestForm" render={() => <EMDClosureRequestForm {...props} />} />
          <Route exact path="/ViewForm" render={() => <ViewForm {...props} />} />


        </Switch>
      </div>
    </div>

  );
}

const Drr: React.FC<ISonaEmdProps> = (props) => {
  return (
    <Router>
      <SonaEmd {...props} />
    </Router>
  );
};

export default Drr;