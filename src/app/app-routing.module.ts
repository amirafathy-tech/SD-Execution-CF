import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { AppComponent } from './app.component';
import { AuthComponent } from './auth/auth.component';
import { AuthGuard } from './auth/auth.guard';
import { ExecutionOrderComponent } from './execution-order/execution-order.component';
import { HomePageComponent } from './home-page/home-page.component';
import { CloudDataComponent } from './cloud-data/cloud-data.component';


const routes: Routes = [
 { path: '', component: HomePageComponent },
  // { path: '', redirectTo: '/home', pathMatch: 'full' }, 
  { path: 'home', component: HomePageComponent },
  { path: 'login', component: AuthComponent },
  { path: 'execution-order',canActivate:[AuthGuard], component: ExecutionOrderComponent },
  { path: 'execution-order-data',canActivate:[AuthGuard], component: CloudDataComponent },
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
