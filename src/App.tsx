import { Switch, Route, Router as WouterRouter } from "wouter";
import CertificatePage from "@/pages/CertificatePage";

function App() {
  return (
    <WouterRouter base={import.meta.env.BASE_URL.replace(/\/$/, "")}>
      <Switch>
        <Route path="/" component={CertificatePage} />
        <Route path="*">
          <div className="min-h-screen flex items-center justify-center">
            <p className="text-muted-foreground">页面不存在</p>
          </div>
        </Route>
      </Switch>
    </WouterRouter>
  );
}

export default App;
